/**
 * ICP-Compare Web 戰略出口合規審計平台 前端邏輯 (App.js)
 */

let allRecords = [];
let currentFilter = "ALL";
let currentSearch = "";
let latestFileName = "";

document.addEventListener("DOMContentLoaded", () => {
    lucide.createIcons();
    initApp();
    setupEventListeners();
});

async function initApp() {
    try {
        const res = await fetch("/api/audit/latest");
        const data = await res.json();

        if (data.status === "success") {
            latestFileName = data.file_name;
            allRecords = data.records || [];
            updateKPIs(data.stats);
            renderCards();
        } else {
            showEmptyState(data.message || "尚未有審計報表資料");
        }
    } catch (err) {
        console.error("載入失敗:", err);
        showEmptyState("連接後端失敗，請確認 web_server.py 是否已啟動");
    }
}

function updateKPIs(stats) {
    if (!stats) return;
    document.getElementById("kpi-total").textContent = stats.total || 0;
    document.getElementById("kpi-rate").textContent = `${stats.auto_release_rate || 0}%`;
    document.getElementById("kpi-high").textContent = stats.high || 0;
    document.getElementById("kpi-medium").textContent = stats.medium || 0;
    document.getElementById("kpi-fp").textContent = stats.fp || 0;

    document.getElementById("count-all").textContent = stats.total || 0;
    document.getElementById("count-high").textContent = stats.high || 0;
    document.getElementById("count-medium").textContent = stats.medium || 0;
    document.getElementById("count-low").textContent = stats.low || 0;
    document.getElementById("count-fp").textContent = stats.fp || 0;
    
    if (stats.file_name) {
        document.getElementById("file-status-text").textContent = `目前載入報表: ${stats.file_name} (${stats.last_updated})`;
    }
}

function renderCards() {
    const container = document.getElementById("cards-container");
    container.innerHTML = "";

    const filtered = allRecords.filter(rec => {
        const level = String(rec["LLM研判等級"] || "").trim();
        const matchesFilter = (currentFilter === "ALL") || (level === currentFilter);

        const searchText = (
            String(rec["查詢名稱"] || "") +
            String(rec["黑名單名稱"] || "") +
            String(rec["條件ID"] || "") +
            String(rec["黑名單ID"] || "") +
            String(rec["查詢地址"] || "") +
            String(rec["黑名單地址"] || "")
        ).toLowerCase();

        const matchesSearch = !currentSearch || searchText.includes(currentSearch.toLowerCase());

        return matchesFilter && matchesSearch;
    });

    if (filtered.length === 0) {
        container.innerHTML = `
            <div class="loading-state">
                <i data-lucide="inbox" style="width: 48px; height: 48px; opacity: 0.5;"></i>
                <p style="margin-top: 12px;">無符合條件的合規案件對照紀錄</p>
            </div>
        `;
        lucide.createIcons();
        return;
    }

    filtered.forEach((rec, idx) => {
        const card = createCardElement(rec, idx + 1);
        container.appendChild(card);
    });

    lucide.createIcons();
}

function getBadgeClass(level) {
    switch (level) {
        case "High": return "badge-high";
        case "Medium": return "badge-medium";
        case "Low": return "badge-low";
        case "False Positive": return "badge-fp";
        default: return "badge-low";
    }
}

function getBadgeLabel(level) {
    switch (level) {
        case "High": return "🔴 High (同一實體/轉運風險)";
        case "Medium": return "🟠 Medium (關聯企業)";
        case "Low": return "🟡 Low (疑慮)";
        case "False Positive": return "🟢 False Positive (確定誤判)";
        default: return level;
    }
}

function createCardElement(rec, index) {
    const card = document.createElement("article");
    card.className = "audit-card glass-card";
    
    const level = String(rec["LLM研判等級"] || "").trim();
    const queryName = rec["查詢名稱"] || "—";
    const watchName = rec["黑名單名稱"] || "—";
    const conditionId = rec["條件ID"] || "—";
    const partyId = rec["黑名單ID"] || "—";
    const pct = rec["原XML命中率"] || "—";
    const riskFactor = rec["風險判定依據"] || "";
    const reasoning = rec["LLM分析推理理由"] || "尚無推理資料";
    const sourceFile = rec["來源檔案"] || "—";

    card.innerHTML = `
        <div class="card-header-bar">
            <div class="case-title">
                <span class="case-num">#${String(index).padStart(2, '0')}</span>
                <span class="entity-pair-name">${escapeHtml(queryName)} <span>⇄</span> ${escapeHtml(watchName)}</span>
            </div>
            <div class="badges-group">
                <span class="pct-badge"><i data-lucide="percent" style="width: 12px; display:inline;"></i> ${pct}</span>
                <span class="badge ${getBadgeClass(level)}">${getBadgeLabel(level)}</span>
            </div>
        </div>

        <div class="comparison-container">
            <div class="entity-box">
                <h4><i data-lucide="user-search"></i> 查詢實體 (Condition)</h4>
                <div class="info-row"><strong>條件ID:</strong> ${escapeHtml(conditionId)}</div>
                <div class="info-row"><strong>名稱:</strong> ${escapeHtml(queryName)}</div>
                <div class="info-row"><strong>國家/城市:</strong> ${escapeHtml(rec["查詢國家"] || "—")} / ${escapeHtml(rec["查詢城市"] || "—")}</div>
                <div class="info-row"><strong>地址:</strong> ${escapeHtml(rec["查詢地址"] || "—")}</div>
            </div>

            <div class="entity-box">
                <h4><i data-lucide="shield-alert"></i> 黑名單限制實體 (Party)</h4>
                <div class="info-row"><strong>黑名單ID:</strong> ${escapeHtml(partyId)}</div>
                <div class="info-row"><strong>黑名單名稱:</strong> ${escapeHtml(watchName)}</div>
                <div class="info-row"><strong>地址:</strong> ${escapeHtml(rec["黑名單地址"] || "—")}</div>
                <div class="info-row"><strong>完整細節:</strong> ${escapeHtml(rec["黑名單完整資訊"] || "—")}</div>
            </div>
        </div>

        ${riskFactor ? `<div class="risk-factor-tag"><i data-lucide="target" style="width: 12px; display:inline;"></i> 風險判定依據：${escapeHtml(riskFactor)}</div>` : ''}

        <div class="reasoning-box">
            <div class="reasoning-header">
                <i data-lucide="brain-circuit"></i> LLM 專家模型 Chain-of-Thought (CoT) 審計推理理由：
            </div>
            <div class="reasoning-body">${escapeHtml(reasoning)}</div>
        </div>

        <div class="card-actions">
            <button class="btn btn-outline btn-action-cert" onclick="openCertModal(${JSON.stringify(rec).replace(/"/g, '&quot;')})">
                <i data-lucide="file-check"></i> 查看稽核憑證
            </button>
            ${level === 'High' ? `<button class="btn btn-action-hold"><i data-lucide="slash"></i> 攔截凍結交易</button>` : ''}
            ${level === 'Medium' ? `<button class="btn btn-action-review"><i data-lucide="file-text"></i> 索取 UBO 證明</button>` : ''}
            ${(level === 'False Positive' || level === 'Low') ? `<button class="btn btn-action-pass"><i data-lucide="check-circle"></i> 自動放行令</button>` : ''}
        </div>
    `;

    return card;
}

function setupEventListeners() {
    // 頁籤過濾
    document.querySelectorAll(".tab-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            currentFilter = btn.getAttribute("data-filter");
            renderCards();
        });
    });

    // 即時搜尋
    const searchInput = document.getElementById("search-input");
    searchInput.addEventListener("input", (e) => {
        currentSearch = e.target.value.trim();
        renderCards();
    });

    // 重新整理
    document.getElementById("refresh-btn").addEventListener("click", initApp);

    // 下載 Excel 報表
    document.getElementById("download-excel-btn").addEventListener("click", () => {
        if (!latestFileName) {
            alert("目前尚無可供下載的 Excel 報表");
            return;
        }
        window.location.href = `/api/reports/download/${latestFileName}`;
    });

    // 啟動審核
    document.getElementById("run-audit-btn").addEventListener("click", startAudit);

    // Modal 開關
    document.getElementById("close-modal-btn").addEventListener("click", closeModal);
    document.getElementById("confirm-modal-btn").addEventListener("click", closeModal);

    // 拖曳上傳
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");

    dropZone.addEventListener("click", () => fileInput.click());
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.style.borderColor = "var(--color-accent-blue)";
    });
    dropZone.addEventListener("dragleave", () => {
        dropZone.style.borderColor = "var(--border-glass-bright)";
    });
    dropZone.addEventListener("drop", async (e) => {
        e.preventDefault();
        dropZone.style.borderColor = "var(--border-glass-bright)";
        if (e.dataTransfer.files.length > 0) {
            await handleFileUpload(e.dataTransfer.files);
        }
    });
    fileInput.addEventListener("change", async () => {
        if (fileInput.files.length > 0) {
            await handleFileUpload(fileInput.files);
        }
    });
}

async function handleFileUpload(files) {
    const formData = new FormData();
    for (let f of files) {
        formData.append("files", f);
    }

    try {
        const res = await fetch("/api/upload", { method: "POST", body: formData });
        const data = await res.json();
        alert(data.message || "上傳完成！點擊「啟動全自動 LLM 合規審核」即可進行比對。");
    } catch (err) {
        alert("上傳失敗: " + err.message);
    }
}

async function startAudit() {
    const btn = document.getElementById("run-audit-btn");
    const progressBox = document.getElementById("progress-box");
    const progressFill = document.getElementById("progress-fill");
    const progressPercent = document.getElementById("progress-percent");

    btn.disabled = true;
    progressBox.classList.remove("hidden");

    try {
        const res = await fetch("/api/audit/run", { method: "POST" });
        const data = await res.json();

        // 輪詢狀態
        const interval = setInterval(async () => {
            const statusRes = await fetch("/api/audit/status");
            const status = await statusRes.json();

            progressFill.style.width = `${status.progress}%`;
            progressPercent.textContent = `${status.progress}%`;

            if (!status.is_running) {
                clearInterval(interval);
                btn.disabled = false;
                progressBox.classList.add("hidden");
                alert("合規審核作業完成！已更新最新數據。");
                await initApp();
            }
        }, 1500);

    } catch (err) {
        alert("啟動審核失敗: " + err.message);
        btn.disabled = false;
        progressBox.classList.add("hidden");
    }
}

function openCertModal(rec) {
    const modal = document.getElementById("cert-modal");
    const content = document.getElementById("modal-content");

    content.innerHTML = `
        <div style="background: rgba(15,23,42,0.8); padding: 20px; border-radius: 10px; border: 1px solid var(--border-glass-bright);">
            <div style="display: flex; justify-content: space-between; border-bottom: 1px solid var(--border-glass); padding-bottom: 12px; margin-bottom: 16px;">
                <div>
                    <h4 style="font-size: 1.1rem; color: #FFF;">合規審計稽核軌跡紀錄</h4>
                    <p style="font-size: 0.8rem; color: var(--text-muted);">Audit Trail Reference ID: ${rec["條件ID"] || 'N/A'}-${rec["黑名單ID"] || 'N/A'}</p>
                </div>
                <div style="text-align: right;">
                    <span style="font-weight: 700; color: var(--color-accent-blue);">${rec["LLM研判等級"]}</span>
                    <p style="font-size: 0.8rem; color: var(--text-muted);">${new Date().toLocaleString()}</p>
                </div>
            </div>

            <div style="margin-bottom: 14px;"><strong>查詢實體：</strong> ${escapeHtml(rec["查詢名稱"])} (${escapeHtml(rec["查詢國家"] || '')})</div>
            <div style="margin-bottom: 14px;"><strong>查詢出貨地址：</strong> ${escapeHtml(rec["查詢地址"])}</div>
            <div style="margin-bottom: 14px;"><strong>命中限制實體：</strong> ${escapeHtml(rec["黑名單名稱"])} (ID: ${escapeHtml(rec["黑名單ID"])})</div>
            <div style="margin-bottom: 14px;"><strong>限制實體物理地址：</strong> ${escapeHtml(rec["黑名單地址"])}</div>
            <div style="margin-bottom: 14px; color: var(--color-medium);"><strong>判定依據：</strong> ${escapeHtml(rec["風險判定依據"] || '無')}</div>
            
            <div style="background: rgba(30,41,59,0.7); padding: 12px; border-radius: 6px; margin-top: 14px;">
                <strong style="color: #60A5FA;">專家 CoT 推理審定結論：</strong>
                <p style="margin-top: 6px; font-size: 0.85rem; line-height: 1.5;">${escapeHtml(rec["LLM分析推理理由"])}</p>
            </div>
        </div>
    `;

    modal.classList.remove("hidden");
    lucide.createIcons();
}

function closeModal() {
    document.getElementById("cert-modal").classList.add("hidden");
}

function showEmptyState(msg) {
    const container = document.getElementById("cards-container");
    container.innerHTML = `
        <div class="loading-state">
            <i data-lucide="alert-circle" style="width: 48px; height: 48px; color: var(--color-medium);"></i>
            <p style="margin-top: 12px;">${escapeHtml(msg)}</p>
        </div>
    `;
    lucide.createIcons();
}

function escapeHtml(text) {
    if (!text) return "";
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
