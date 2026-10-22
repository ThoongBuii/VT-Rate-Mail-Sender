let state = null;
let pollTimer = null;
let currentView = "compose";
let previewIndex = 0;
let saveTimer = null;
let axPromptedThisSession = false;

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

function composeEl() {
  return document.getElementById("composeEditor");
}

const DEFAULT_FONT = "Aptos";
const DEFAULT_SIZE_PT = "12";
const OUTLOOK_FONT_STACK = 'Aptos, Calibri, Arial, sans-serif';

function setComposeHtml(html) {
  composeEl().innerHTML = html || "";
}

function getComposeHtml({ normalize = true } = {}) {
  const ed = composeEl();
  let html = ed.innerHTML || "";
  if (!html.replace(/<br\s*\/?>/gi, "").replace(/&nbsp;/gi, "").trim()) return "";
  // Autosave: bỏ normalize (DOM parse bảng lớn → giật). Chỉ normalize lúc Preview/Send.
  if (!normalize) return html;
  return normalizeHtmlForOutlook(html);
}

/** Chuẩn hóa nhẹ cho Outlook — KHÔNG ép đè font user đã chọn trên toolbar. */
function normalizeHtmlForOutlook(html) {
  const trimmed = (html || "").trim();
  if (!trimmed) return "";

  const box = document.createElement("div");
  box.innerHTML = trimmed;

  box.querySelectorAll("table").forEach((t) => {
    t.style.borderCollapse = t.style.borderCollapse || "collapse";
    t.style.borderSpacing = t.style.borderSpacing || "0";
    if (!t.getAttribute("border")) t.setAttribute("border", "1");
    if (!t.getAttribute("cellspacing")) t.setAttribute("cellspacing", "0");
    if (!t.getAttribute("cellpadding")) t.setAttribute("cellpadding", "4");
  });

  // px → pt trên style inline (giữ nguyên font-family user chọn)
  box.querySelectorAll("[style]").forEach((el) => {
    const s = el.getAttribute("style") || "";
    const next = s.replace(/font-size\s*:\s*(\d+(?:\.\d+)?)px/gi, (_, px) => {
      const pt = Math.round((Number(px) * 72) / 96 * 10) / 10;
      return `font-size:${pt}pt`;
    });
    if (next !== s) el.setAttribute("style", next);
  });

  // Ảnh: bỏ max-width:100% (Outlook hay thu nhỏ); giữ width/height px nếu có
  box.querySelectorAll("img").forEach((img) => {
    let style = img.getAttribute("style") || "";
    style = style.replace(/max-width\s*:\s*100%\s*;?/gi, "").trim();
    const wAttr = img.getAttribute("width");
    const hAttr = img.getAttribute("height");
    if (wAttr && !/width\s*:/i.test(style)) style = `${style};width:${wAttr}px`.replace(/^;/, "");
    if (hAttr && !/height\s*:/i.test(style)) style = `${style};height:${hAttr}px`.replace(/^;/, "");
    if (style) img.setAttribute("style", style);
    else img.removeAttribute("style");
  });

  let root = box.firstElementChild;
  if (root && root.tagName === "DIV" && box.childNodes.length === 1) {
    root.setAttribute("data-vt-body", "1");
    root.style.margin = root.style.margin || "0";
    root.style.padding = root.style.padding || "0";
    root.style.lineHeight = root.style.lineHeight || "1.35";
    return box.innerHTML;
  }

  const wrap = document.createElement("div");
  wrap.setAttribute("data-vt-body", "1");
  wrap.setAttribute(
    "style",
    `line-height:1.35;margin:0;padding:0;color:#222;`
  );
  while (box.firstChild) wrap.appendChild(box.firstChild);
  return wrap.outerHTML;
}

function scheduleSaveCompose() {
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    saveCompose({ normalize: false }).catch(() => {});
  }, 1200);
}

function frameDoc(html) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      html,body{margin:0;padding:12px;background:#fff;color:#222}
      body{font-family:Aptos,Calibri,"Segoe UI",Arial,sans-serif;font-size:12pt;line-height:1.35}
      table{border-collapse:collapse}
      img{max-width:100%;height:auto}
      p{margin:0 0 0.6em}
    </style></head><body>${html || ""}</body></html>`;
}

function focusEditor() {
  const ed = composeEl();
  ed.focus();
}

/** Lưu/khôi phục selection — dropdown toolbar hay làm mất vùng bôi đen. */
let savedEditorRange = null;

function saveEditorSelection() {
  const ed = composeEl();
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  if (!ed.contains(range.commonAncestorContainer)) return;
  savedEditorRange = range.cloneRange();
}

function restoreEditorSelection() {
  const ed = composeEl();
  ed.focus();
  if (!savedEditorRange) return false;
  const sel = window.getSelection();
  sel.removeAllRanges();
  try {
    sel.addRange(savedEditorRange);
    return true;
  } catch (_) {
    return false;
  }
}

function rangeIntersectsNode(range, node) {
  try {
    if (typeof range.intersectsNode === "function") {
      return range.intersectsNode(node);
    }
  } catch (_) {}
  const nodeRange = document.createRange();
  try {
    nodeRange.selectNode(node);
  } catch (_) {
    nodeRange.selectNodeContents(node);
  }
  return (
    range.compareBoundaryPoints(Range.END_TO_START, nodeRange) < 0 &&
    range.compareBoundaryPoints(Range.START_TO_END, nodeRange) > 0
  );
}

/**
 * Áp style lên TOÀN BỘ vùng bôi đen (mọi thẻ con: p/span/td/font…).
 * Không dựa vào execCommand(fontName) — hay bỏ sót Dear / ô bảng.
 */
function applyStyleToSelection(styleMap) {
  const ed = composeEl();
  restoreEditorSelection();
  focusEditor();
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;

  let range = sel.getRangeAt(0);
  if (range.collapsed && savedEditorRange && !savedEditorRange.collapsed) {
    sel.removeAllRanges();
    sel.addRange(savedEditorRange);
    range = sel.getRangeAt(0);
  }

  if (range.collapsed) {
    const parts = [];
    if (styleMap.fontFamily) parts.push(`font-family:${styleMap.fontFamily}`);
    if (styleMap.fontSize) parts.push(`font-size:${styleMap.fontSize}`);
    if (styleMap.color) parts.push(`color:${styleMap.color}`);
    if (styleMap.backgroundColor) parts.push(`background-color:${styleMap.backgroundColor}`);
    if (parts.length) wrapSelectionWithSpan(parts.join(";"));
    return;
  }

  const targets = [];
  const walker = document.createTreeWalker(ed, NodeFilter.SHOW_ELEMENT, null);
  let node = walker.nextNode();
  while (node) {
    if (rangeIntersectsNode(range, node)) {
      const tag = node.tagName;
      if (tag !== "TABLE" && tag !== "TBODY" && tag !== "THEAD" && tag !== "TR" && tag !== "COLGROUP") {
        targets.push(node);
      }
    }
    node = walker.nextNode();
  }

  // Nếu không bắt được thẻ (chỉ text thuần) → bọc span
  if (!targets.length) {
    const parts = [];
    if (styleMap.fontFamily) parts.push(`font-family:${styleMap.fontFamily}`);
    if (styleMap.fontSize) parts.push(`font-size:${styleMap.fontSize}`);
    if (styleMap.color) parts.push(`color:${styleMap.color}`);
    if (styleMap.backgroundColor) parts.push(`background-color:${styleMap.backgroundColor}`);
    wrapSelectionWithSpan(parts.join(";"));
    return;
  }

  targets.forEach((el) => {
    if (styleMap.fontFamily) {
      el.style.fontFamily = styleMap.fontFamily;
      if (el.tagName === "FONT") el.removeAttribute("face");
    }
    if (styleMap.fontSize) {
      el.style.fontSize = styleMap.fontSize;
      if (el.tagName === "FONT") el.removeAttribute("size");
    }
    if (styleMap.color) {
      el.style.color = styleMap.color;
      if (el.tagName === "FONT") el.removeAttribute("color");
    }
    if (styleMap.backgroundColor) {
      el.style.backgroundColor = styleMap.backgroundColor;
    }
  });

  scheduleSaveCompose();
  syncFormatToolbar();
  saveEditorSelection();
}

function runFormat(cmd, value = null) {
  restoreEditorSelection();
  focusEditor();
  try {
    document.execCommand("styleWithCSS", false, true);
  } catch (_) {}
  document.execCommand(cmd, false, value);
  scheduleSaveCompose();
  syncFormatToolbar();
  saveEditorSelection();
}

function wrapSelectionWithSpan(styleText) {
  restoreEditorSelection();
  focusEditor();
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  if (range.collapsed) {
    const span = document.createElement("span");
    span.setAttribute("style", styleText);
    span.appendChild(document.createTextNode("\u200b"));
    range.insertNode(span);
    range.setStart(span.firstChild, 1);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    scheduleSaveCompose();
    saveEditorSelection();
    return;
  }
  try {
    const span = document.createElement("span");
    span.setAttribute("style", styleText);
    range.surroundContents(span);
  } catch (_) {
    const frag = range.extractContents();
    const span = document.createElement("span");
    span.setAttribute("style", styleText);
    span.appendChild(frag);
    range.insertNode(span);
  }
  scheduleSaveCompose();
  syncFormatToolbar();
  saveEditorSelection();
}

function applyFontFamily(name) {
  const font = (name || DEFAULT_FONT).trim();
  // Stack nhẹ để Outlook fallback nếu máy thiếu font
  const stack = font.includes(",") ? font : `${font}, Calibri, Arial, sans-serif`;
  applyStyleToSelection({ fontFamily: stack });
}

function applyFontSizePt(pt) {
  const size = `${String(pt || DEFAULT_SIZE_PT)}pt`;
  applyStyleToSelection({ fontSize: size });
}

function applyForeColor(color) {
  applyStyleToSelection({ color });
}

function applyHiliteColor(color) {
  applyStyleToSelection({ backgroundColor: color });
}

function syncFormatToolbar() {
  const map = {
    bold: "bold",
    italic: "italic",
    underline: "underline",
    justifyLeft: "justifyLeft",
    justifyCenter: "justifyCenter",
    justifyRight: "justifyRight",
    justifyFull: "justifyFull",
  };
  document.querySelectorAll(".fmt-btn[data-fmt]").forEach((btn) => {
    const cmd = btn.getAttribute("data-fmt");
    if (!map[cmd] && cmd !== "bold" && cmd !== "italic" && cmd !== "underline") {
      btn.classList.remove("active");
      return;
    }
    if (["bold", "italic", "underline", "justifyLeft", "justifyCenter", "justifyRight", "justifyFull"].includes(cmd)) {
      try {
        btn.classList.toggle("active", document.queryCommandState(cmd));
      } catch (_) {
        btn.classList.remove("active");
      }
    }
  });
}

function initFormatToolbar() {
  const bar = document.getElementById("formatBar");
  if (!bar) return;

  // Giữ selection khi click toolbar (kể cả lúc mở dropdown)
  bar.addEventListener("mousedown", (ev) => {
    saveEditorSelection();
    if (ev.target.closest("select, input")) return;
    ev.preventDefault();
  });

  bar.querySelectorAll(".fmt-btn[data-fmt]").forEach((btn) => {
    btn.addEventListener("click", () => {
      restoreEditorSelection();
      const cmd = btn.getAttribute("data-fmt");
      runFormat(cmd);
    });
  });

  document.getElementById("fmtFont").addEventListener("focus", saveEditorSelection);
  document.getElementById("fmtSize").addEventListener("focus", saveEditorSelection);
  document.getElementById("fmtFont").addEventListener("change", (ev) => {
    applyFontFamily(ev.target.value);
  });
  document.getElementById("fmtSize").addEventListener("change", (ev) => {
    applyFontSizePt(ev.target.value);
  });
  document.getElementById("fmtForeColor").addEventListener("focus", saveEditorSelection);
  document.getElementById("fmtHilite").addEventListener("focus", saveEditorSelection);
  document.getElementById("fmtForeColor").addEventListener("input", (ev) => {
    applyForeColor(ev.target.value);
  });
  document.getElementById("fmtHilite").addEventListener("input", (ev) => {
    applyHiliteColor(ev.target.value);
  });

  const ed = composeEl();
  ed.addEventListener("keyup", () => {
    saveEditorSelection();
    syncFormatToolbar();
  });
  ed.addEventListener("mouseup", () => {
    saveEditorSelection();
    syncFormatToolbar();
  });
  ed.addEventListener("focus", syncFormatToolbar);
  document.addEventListener("selectionchange", () => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    if (ed.contains(sel.anchorNode)) saveEditorSelection();
  });
}

function insertHtmlAtCursor(html) {
  const ed = composeEl();
  ed.focus();
  // insertHTML hỗ trợ Ctrl+Z của trình duyệt / WebView
  const ok = document.execCommand("insertHTML", false, html);
  if (!ok) {
    // Fallback Selection API
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      const tip = range.createContextualFragment(html);
      range.insertNode(tip);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      ed.insertAdjacentHTML("beforeend", html);
    }
  }
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

function renderCountrySelect() {
  const sel = document.getElementById("countrySelect");
  const hint = document.getElementById("libraryHint");
  if (!sel) return;
  const lib = state?.library || {};
  const countries = lib.countries || [];
  const active = state?.active_country || lib.active_country || "";
  const prev = sel.value;
  sel.innerHTML = "";
  if (!countries.length) {
    sel.innerHTML = `<option value="">— Chưa có —</option>`;
  } else {
    countries.forEach((c) => {
      const n = (lib.counts && lib.counts[c]) || "";
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = n ? `${c} (${n})` : c;
      if (c === active) opt.selected = true;
      sel.appendChild(opt);
    });
    if (!active && prev && countries.includes(prev)) sel.value = prev;
  }
  if (hint) {
    hint.textContent = lib.source_name
      ? `Thư viện: ${lib.source_name}`
      : "Nạp Excel 1 lần (mỗi sheet = 1 country) → chọn Country.";
  }
}

let dragFromIndex = null;

function renderList() {
  const box = document.getElementById("mailList");
  box.innerHTML = "";
  renderCountrySelect();
  if (!state || !state.mails.length) {
    box.innerHTML = `<div class="mail-item"><div class="meta"><div class="company">Chưa có agency — chọn Country hoặc + Thêm</div></div></div>`;
    return;
  }
  state.mails.forEach((m, i) => {
    const active =
      currentView === "preview" ? i === previewIndex : i === state.selected_index;
    const el = document.createElement("div");
    el.className = "mail-item" + (active ? " active" : "");
    el.draggable = true;
    el.dataset.index = String(i);
    el.innerHTML = `
      <div class="drag" title="Kéo thả đổi thứ tự">⋮⋮</div>
      <div class="idx">${i + 1}</div>
      <div class="meta">
        <div class="company">${escapeHtml(m.agency_company || "(no company)")}</div>
        <div class="email">${escapeHtml(m.account_mail || "")}</div>
      </div>
      <div class="right">
        <div class="st st-${m.status}">${m.status}</div>
        <button type="button" class="btn-del" title="Xóa" data-del="${i}">×</button>
      </div>`;
    el.onclick = (ev) => {
      if (ev.target.closest("[data-del]") || ev.target.closest(".drag")) return;
      selectMail(i);
    };
    el.querySelector("[data-del]").onclick = async (ev) => {
      ev.stopPropagation();
      if (!confirm(`Xóa agency?\n${m.agency_company || ""}\n${m.account_mail || ""}`)) return;
      try {
        const res = await api("/api/mails/delete", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ index: i }),
        });
        applyState(res.state);
        if (currentView === "preview") await loadPreview(Math.min(i, Math.max(0, res.state.mails.length - 1)));
      } catch (e) {
        alert(e.message);
      }
    };
    el.addEventListener("dragstart", (ev) => {
      dragFromIndex = i;
      el.classList.add("dragging");
      try {
        ev.dataTransfer.setData("text/plain", String(i));
        ev.dataTransfer.effectAllowed = "move";
      } catch (_) {}
    });
    el.addEventListener("dragend", () => {
      el.classList.remove("dragging");
      dragFromIndex = null;
    });
    el.addEventListener("dragover", (ev) => {
      ev.preventDefault();
      ev.dataTransfer.dropEffect = "move";
    });
    el.addEventListener("drop", async (ev) => {
      ev.preventDefault();
      const from = dragFromIndex;
      const to = i;
      if (from == null || from === to) return;
      const order = state.mails.map((_, idx) => idx);
      const [moved] = order.splice(from, 1);
      order.splice(to, 0, moved);
      try {
        const res = await api("/api/mails/reorder", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ order }),
        });
        applyState(res.state);
        previewIndex = to;
        if (currentView === "preview") await loadPreview(to);
      } catch (e) {
        alert(e.message);
      }
    });
    box.appendChild(el);
  });
}

function attachmentItemsFromState(s) {
  if (Array.isArray(s?.attachments) && s.attachments.length) {
    return s.attachments.map((p) => ({
      path: p,
      name: String(p).split(/[/\\]/).pop(),
    }));
  }
  const raw = (s?.attachment || "").trim();
  if (!raw) return [];
  return raw.split("|||").filter(Boolean).map((p) => ({
    path: p.trim(),
    name: p.trim().split(/[/\\]/).pop(),
  }));
}

function setAttachmentUi(itemsOrPath) {
  const box = document.getElementById("attachList");
  if (!box) return;
  let items = [];
  if (Array.isArray(itemsOrPath)) {
    items = itemsOrPath.map((it) =>
      typeof it === "string"
        ? { path: it, name: String(it).split(/[/\\]/).pop() }
        : { path: it.path, name: it.name || String(it.path).split(/[/\\]/).pop() }
    );
  } else if (itemsOrPath) {
    items = attachmentItemsFromState({ attachment: String(itemsOrPath) });
  }
  if (state) {
    state.attachment = items.map((i) => i.path).join("|||");
    state.attachments = items.map((i) => i.path);
  }
  if (!items.length) {
    box.innerHTML = `<span class="attach-empty">Chưa chọn file</span>`;
    return;
  }
  box.innerHTML = items
    .map(
      (it) => `
      <span class="attach-chip" title="${escapeHtml(it.path)}">
        <span>${escapeHtml(it.name)}</span>
        <button type="button" data-path="${escapeHtml(it.path)}" aria-label="Xóa">×</button>
      </span>`
    )
    .join("");
  box.querySelectorAll("button[data-path]").forEach((btn) => {
    btn.onclick = async () => {
      try {
        const res = await api(
          `/api/attachment?path=${encodeURIComponent(btn.getAttribute("data-path") || "")}`,
          { method: "DELETE" }
        );
        setAttachmentUi(res.attachments || res.paths || []);
      } catch (e) {
        alert(e.message);
      }
    };
  });
}

async function uploadAttachmentFiles(fileList) {
  const files = Array.from(fileList || []).filter(Boolean);
  if (!files.length) return;
  const fd = new FormData();
  files.forEach((f) => fd.append("file", f));
  const res = await api("/api/attachment", { method: "POST", body: fd });
  setAttachmentUi(res.attachments || res.paths || []);
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Không đọc được ảnh"));
    reader.readAsDataURL(file);
  });
}

function loadImageSize(uri) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || 0, h: img.naturalHeight || 0 });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = uri;
  });
}

async function insertImagesFromFiles(files) {
  const MAX_W = 650;
  for (const file of files) {
    if (!file.type.startsWith("image/")) continue;
    const uri = await readFileAsDataUrl(file);
    if (!uri) continue;
    const { w, h } = await loadImageSize(uri);
    let width = w;
    let height = h;
    if (width > MAX_W && width > 0) {
      height = Math.max(1, Math.round((MAX_W * height) / width));
      width = MAX_W;
    }
    const sizeAttrs =
      width && height
        ? ` width="${width}" height="${height}" style="width:${width}px;height:${height}px;"`
        : ` style="height:auto;"`;
    insertHtmlAtCursor(
      `<p style="margin:0 0 0.6em;"><img src="${uri}" alt="${escapeHtml(file.name || "image")}"${sizeAttrs} /></p>`
    );
  }
  scheduleSaveCompose();
}

function bindDropZone(el, { onFiles }) {
  if (!el) return;
  el.addEventListener("dragenter", (ev) => {
    ev.preventDefault();
    el.classList.add("drag-over");
  });
  el.addEventListener("dragover", (ev) => {
    ev.preventDefault();
    el.classList.add("drag-over");
  });
  el.addEventListener("dragleave", () => el.classList.remove("drag-over"));
  el.addEventListener("drop", async (ev) => {
    ev.preventDefault();
    el.classList.remove("drag-over");
    const files = Array.from(ev.dataTransfer?.files || []);
    if (!files.length) return;
    try {
      await onFiles(files);
    } catch (e) {
      alert(e.message || String(e));
    }
  });
}

function updateAxBanner(s) {
  const banner = document.getElementById("axBanner");
  if (!banner) return;
  const need =
    s.platform === "Darwin" &&
    s.accessibility_needed !== false &&
    !s.accessibility_trusted;
  banner.hidden = !need;
  const text = document.getElementById("axBannerText");
  if (text && need) {
    text.textContent =
      s.accessibility_message ||
      "Cần bật Accessibility cho «VT Rate Mail Sender» để gửi mail mượt (paste body).";
  }
}

async function ensureMacAccessibilityOnce(s) {
  // Chữ ký do user dán — không bắt Accessibility
  if (!s || s.platform !== "Darwin") return;
  updateAxBanner({ ...s, accessibility_needed: false });
  axPromptedThisSession = true;
}

function applyState(s, forceHtml = false) {
  state = s;
  document.getElementById("subject").value = s.subject || "";
  setAttachmentUi(attachmentItemsFromState(s));
  document.getElementById("delayMin").value = s.delay_min;
  document.getElementById("delayMax").value = s.delay_max;
  document.getElementById("stats").textContent =
    `${s.stats.total} agency · ready ${s.stats.ready} · sent ${s.stats.sent} · fail ${s.stats.failed}`;
  updateAxBanner(s);
  if (s.outlook_ready) {
    setOutlookStatus(`Outlook sẵn sàng · ${s.outlook_account || ""}`);
  } else {
    setOutlookStatus("Outlook: chưa kết nối");
  }
  const btnSig = document.getElementById("btnCaptureSig");
  if (btnSig) {
    // Mac: nút dự phòng (tuỳ chọn). Windows: ẩn — chữ ký New Mail tự gắn.
    const isMac = s.platform === "Darwin";
    btnSig.hidden = !isMac;
    if (isMac) {
      const cached = !!(s.mac_signature_cached || s.mac_signature_has_file);
      const ready = !!s.mac_signature_ready;
      btnSig.textContent = ready ? "Chữ ký ✓" : "Chữ ký";
      btnSig.title =
        s.mac_signature_message ||
        "Dán chữ ký Outlook vào app (1 lần) rồi Semi-Auto";
      btnSig.classList.toggle("ok-sig", ready);
    }
  }
  if (forceHtml || !getComposeHtml({ normalize: false })) {
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
  const normalize = extra.normalize !== false;
  const { normalize: _n, ...rest } = extra;
  const payload = {
    subject: document.getElementById("subject").value,
    template_html: getComposeHtml({ normalize }),
    delay_min: Number(document.getElementById("delayMin").value || 10),
    delay_max: Number(document.getElementById("delayMax").value || 20),
    selected_index: state ? state.selected_index : 0,
    ...rest,
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
  const tick = async () => {
    try {
      // Đang soạn / mở dialog chữ ký → poll thưa, tránh giật UI
      const sigOpen = document.getElementById("sigDialog")?.open;
      const edFocused = document.activeElement === composeEl();
      const running = !!state?.progress?.is_running;
      if (!running && (sigOpen || edFocused)) {
        pollTimer = setTimeout(tick, 4000);
        return;
      }
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
    pollTimer = setTimeout(tick, state?.progress?.is_running ? 900 : 2500);
  };
  tick();
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
    await refreshState();
    alert(res.message || "Outlook sẵn sàng");
  } catch (e) {
    alert(e.message);
  }
};

/** HTML chữ ký đã nhúng ảnh (từ Clipboard) — không lấy từ contenteditable (dễ mất ảnh). */
let pendingSigHtml = "";

function sigPasteBox() {
  return document.getElementById("sigPasteBox");
}

async function openSigDialog() {
  const dlg = document.getElementById("sigDialog");
  const box = sigPasteBox();
  if (!dlg || !box) return;
  pendingSigHtml = "";
  box.innerHTML = "";
  try {
    const prev = await api("/api/outlook/signature-preview");
    if (prev.preview_html) {
      pendingSigHtml = prev.preview_html;
      box.innerHTML = prev.preview_html;
    }
  } catch (_) {
    /* empty */
  }
  if (typeof dlg.showModal === "function") dlg.showModal();
  else dlg.setAttribute("open", "");
}

document.getElementById("btnCaptureSig").onclick = () => {
  openSigDialog().catch((e) => alert(e.message || String(e)));
};

document.getElementById("btnSigCancel")?.addEventListener("click", () => {
  document.getElementById("sigDialog")?.close();
});

document.getElementById("btnSigClear")?.addEventListener("click", () => {
  pendingSigHtml = "";
  const box = sigPasteBox();
  if (box) box.innerHTML = "";
});

document.getElementById("btnSigFromClip")?.addEventListener("click", async () => {
  try {
    // Preview đã nhúng data URI — không dùng Cmd+V vào khung (mất ảnh)
    const res = await api("/api/outlook/preview-signature-clipboard", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    });
    pendingSigHtml = res.preview_html || "";
    const box = sigPasteBox();
    if (box) box.innerHTML = pendingSigHtml;
    if (!res.image_count) {
      alert(
        "Clipboard chưa có ảnh logo.\n" +
          "Trong Outlook chọn cả chữ ký + logo → Cmd+C → bấm lại «Dán từ Clipboard»."
      );
    }
  } catch (e) {
    alert(e.message || String(e));
  }
});

document.getElementById("btnSigSave")?.addEventListener("click", async () => {
  try {
    // Luôn gửi HTML đã chuẩn bị (pending) hoặc đọc lại clipboard — không tin contenteditable
    const html = pendingSigHtml || "";
    const res = await api("/api/outlook/save-signature", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ html }),
    });
    pendingSigHtml = res.preview_html || pendingSigHtml;
    const box = sigPasteBox();
    if (box && res.preview_html) box.innerHTML = res.preview_html;
    await refreshState();
    alert(res.message || "Đã lưu chữ ký");
    document.getElementById("sigDialog")?.close();
  } catch (e) {
    alert(e.message || String(e));
  }
});

const editor = composeEl();
editor.addEventListener("input", scheduleSaveCompose);
editor.addEventListener("blur", () => {
  saveCompose().catch(() => {});
});
editor.addEventListener("paste", async (ev) => {
  ev.preventDefault();
  ev.stopPropagation();

  // Ảnh thuần từ clipboard (Win/Mac) → data URI trong body
  const imageFiles = [];
  try {
    const items = Array.from(ev.clipboardData?.items || []);
    for (const it of items) {
      if (it.kind === "file" && it.type.startsWith("image/")) {
        const f = it.getAsFile();
        if (f) imageFiles.push(f);
      }
    }
  } catch (_) {}
  if (imageFiles.length) {
    await insertImagesFromFiles(imageFiles);
    return;
  }

  let browserHtml = "";
  try {
    browserHtml = ev.clipboardData?.getData("text/html") || "";
  } catch (_) {}
  try {
    const res = await api("/api/clipboard/paste", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ html: browserHtml || "", replace: false }),
    });
    insertHtmlAtCursor(res.html || res.template_html || "");
    scheduleSaveCompose();
  } catch (e) {
    const text = ev.clipboardData?.getData("text/plain") || "";
    if (text) {
      document.execCommand("insertText", false, text);
      scheduleSaveCompose();
    } else {
      alert(e.message);
    }
  }
});

bindDropZone(editor, {
  async onFiles(files) {
    const images = files.filter((f) => f.type.startsWith("image/"));
    const others = files.filter((f) => !f.type.startsWith("image/"));
    if (images.length) await insertImagesFromFiles(images);
    if (others.length) await uploadAttachmentFiles(others);
  },
});

// Preview: kéo thả ảnh → soạn; file khác → đính kèm (đồng bộ Soạn)
bindDropZone(document.querySelector(".preview-main"), {
  async onFiles(files) {
    const images = files.filter((f) => f.type.startsWith("image/"));
    const others = files.filter((f) => !f.type.startsWith("image/"));
    if (images.length) {
      await insertImagesFromFiles(images);
      await saveCompose();
      if (currentView === "preview") await loadPreview(previewIndex);
    }
    if (others.length) {
      await uploadAttachmentFiles(others);
      if (currentView === "preview") await loadPreview(previewIndex);
    }
  },
});
bindDropZone(document.querySelector(".attach-block"), {
  async onFiles(files) {
    await uploadAttachmentFiles(files);
  },
});

document.getElementById("btnClearAttach").onclick = async () => {
  try {
    await api("/api/attachment", { method: "DELETE" });
    setAttachmentUi([]);
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
    const countries = (res.countries || []).join(", ") || res.active_country || "";
    alert(
      `Đã lưu thư viện agency.\n` +
        `Country: ${countries}\n` +
        `Đang mở: ${res.active_country || ""} · ${res.count} agency`
    );
  } catch (e) {
    alert(e.message);
  } finally {
    ev.target.value = "";
  }
};

document.getElementById("countrySelect").onchange = async (ev) => {
  const country = ev.target.value;
  if (!country) return;
  try {
    await saveCompose();
    const res = await api("/api/library/select", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ country }),
    });
    previewIndex = 0;
    applyState(res.state, true);
    if (currentView === "preview") await loadPreview(0);
  } catch (e) {
    alert(e.message);
    renderCountrySelect();
  }
};

document.getElementById("btnAddAgency").onclick = () => {
  const dlg = document.getElementById("agencyDialog");
  document.getElementById("agencyDialogTitle").textContent = "Thêm agency";
  document.getElementById("dlgCompany").value = "";
  document.getElementById("dlgName").value = "";
  document.getElementById("dlgMail").value = "";
  document.getElementById("dlgCc").value = "";
  if (typeof dlg.showModal === "function") dlg.showModal();
  else alert("Trình duyệt không hỗ trợ dialog.");
};

document.getElementById("btnAgencyCancel").onclick = () => {
  document.getElementById("agencyDialog").close();
};

document.getElementById("agencyForm").addEventListener("submit", async (ev) => {
  ev.preventDefault();
  const dlg = document.getElementById("agencyDialog");
  try {
    const res = await api("/api/mails/add", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        agency_company: document.getElementById("dlgCompany").value,
        account_name: document.getElementById("dlgName").value,
        account_mail: document.getElementById("dlgMail").value,
        mail_cc: document.getElementById("dlgCc").value,
      }),
    });
    applyState(res.state);
    previewIndex = res.index ?? res.state.selected_index ?? 0;
    if (currentView === "preview") await loadPreview(previewIndex);
    dlg.close();
  } catch (e) {
    alert(e.message);
  }
});

document.getElementById("fileAttach").onchange = async (ev) => {
  try {
    await uploadAttachmentFiles(ev.target.files);
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
      alert("Chưa có nội dung mail. Hãy soạn hoặc Ctrl+V từ Outlook vào khung soạn.");
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
    if (state?.progress?.is_paused) {
      // Lưu SOẠN trước Resume — các mail còn lại dùng nội dung vừa sửa.
      await saveCompose();
      await api("/api/send/resume", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          subject: document.getElementById("subject").value,
          template_html: getComposeHtml(),
          attachment: state?.attachment || "",
        }),
      });
    } else {
      await api("/api/send/pause", { method: "POST" });
    }
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

document.getElementById("btnAxOpen")?.addEventListener("click", async () => {
  try {
    const st = await api("/api/accessibility/ensure", { method: "POST" });
    if (state) {
      state.accessibility_trusted = !!st.trusted;
      state.accessibility_message = st.message || "";
      updateAxBanner(state);
    }
  } catch (e) {
    alert(e.message || String(e));
  }
});
document.getElementById("btnAxDismiss")?.addEventListener("click", () => {
  const banner = document.getElementById("axBanner");
  if (banner) banner.hidden = true;
});

(async function boot() {
  initFormatToolbar();
  const s = await api("/api/state");
  applyState(s, true);
  await ensureMacAccessibilityOnce(s);
  startPolling();
})();
