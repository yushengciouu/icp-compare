/**
 * ICP-Compare Web 戰略出口合規審計平台 前端邏輯 (App.js)
 * 支援多使用者 Session / Job-ID 獨立隔離與頁面清空重置
 */

let allRecords = [];
let currentFilter = "ALL";
let currentSearch = "";
let latestFileName = "";
let currentJobId = sessionStorage.getItem("currentJobId") || null;

document.addEventListener("DOMContentLoaded", () => {
    lucide.createIcons();
    initApp();
    setupEventListeners();
});

async function initApp() {
    if (!currentJobId) {
        // 初始狀態或已清空：顯示純淨待上傳狀態
        resetPageToInitialState();
        return;
    }

    try {
        const res = await fetch(`/api/audit/results?job_id=${currentJobId}`);
        const data = await res.json();

        if (data.status === "success" && data.records !== undefined) {
            latestFileName = data.stats ? (data.stats.file_name || "") : "";
            allRecords = data.records || [];
            updateKPIs(data.stats);
            if (allRecords.length === 0) {
                showEmptyState("✅ 審算完成：上傳之 XML 報文全數合格（無命中率 >= 75% 之疑慮對照實體）");
            } else {
                renderCards();
            }
        } else {
            resetPageToInitialState();
        }
    } catch (err) {
        console.error("載入失敗:", err);
        resetPageToInitialState();
    }
}

function resetPageToInitialState() {
    currentJobId = null;
    sessionStorage.removeItem("currentJobId");
    allRecords = [];
    latestFileName = "";

    updateKPIs({
        total: 0,
        high: 0,
        medium: 0,
        low: 0,
        fp: 0,
        auto_release_rate: 0.0,
        file_name: "",
        last_updated: ""
    });

    document.getElementById("file-status-text").textContent = "目前狀態: 乾淨初始狀態 (等待上傳 XML 報文檔案)";
    showEmptyState("當前畫面已清空，請上傳 XML 報文檔案開始進行合規審算");
}

function updateKPIs(stats) {
    if (!stats) return;
    const highCount = Number(stats.high || 0);
    const mediumCount = Number(stats.medium || 0);

    document.getElementById("kpi-total").textContent = stats.total || 0;
    document.getElementById("kpi-rate").textContent = `${stats.auto_release_rate || 0}%`;
    document.getElementById("kpi-high").textContent = highCount;
    document.getElementById("kpi-medium").textContent = mediumCount;
    document.getElementById("kpi-fp").textContent = stats.fp || 0;

    // 動態警示發光邏輯：僅在非 0 時觸發發光
    const highCard = document.getElementById("kpi-card-high");
    if (highCard) {
        if (highCount > 0) {
            highCard.classList.add("has-warning-high");
        } else {
            highCard.classList.remove("has-warning-high");
        }
    }

    const mediumCard = document.getElementById("kpi-card-medium");
    if (mediumCard) {
        if (mediumCount > 0) {
            mediumCard.classList.add("has-warning-medium");
        } else {
            mediumCard.classList.remove("has-warning-medium");
        }
    }

    document.getElementById("count-all").textContent = stats.total || 0;
    document.getElementById("count-high").textContent = highCount;
    document.getElementById("count-medium").textContent = mediumCount;
    document.getElementById("count-low").textContent = stats.low || 0;
    document.getElementById("count-fp").textContent = stats.fp || 0;
    
    if (stats.file_name) {
        document.getElementById("file-status-text").textContent = `任務備忘: ${stats.file_name} (${stats.last_updated})`;
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
    const nameMatch = rec["公司名稱比對"] || "—";
    const addressMatch = rec["地址比對"] || "—";
    const reasoning = rec["LLM分析推理理由"] || "尚無推理資料";

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

        <div class="risk-factor-tag"><i data-lucide="target" style="width: 12px; display:inline;"></i> 公司名稱：${escapeHtml(nameMatch)} ｜ 地址比對：${escapeHtml(addressMatch)}</div>

        <div class="reasoning-box">
            <div class="reasoning-header">
                <i data-lucide="brain-circuit"></i> LLM 專家模型 Chain-of-Thought (CoT) 審計推理理由：
            </div>
            <div class="reasoning-body">${escapeHtml(reasoning)}</div>
        </div>

        <div class="card-footer-actions">
            ${level === 'High' ? `
                <div class="suggestion-tag suggestion-hold">
                    <i data-lucide="slash" style="width:15px;height:15px;"></i>
                    <span>建議處置：<strong>攔截凍結交易</strong></span>
                </div>
            ` : ''}
            ${level === 'Medium' ? `
                <div class="suggestion-tag suggestion-review">
                    <i data-lucide="file-text" style="width:15px;height:15px;"></i>
                    <span>建議處置：<strong>人工二審</strong></span>
                </div>
            ` : ''}
            ${(level === 'False Positive' || level === 'Low') ? `
                <div class="suggestion-tag suggestion-pass">
                    <i data-lucide="check-circle" style="width:15px;height:15px;"></i>
                    <span>建議處置：<strong>自動放行</strong></span>
                </div>
            ` : ''}
        </div>
    `;

    return card;
}

function setupEventListeners() {
    // 清空與重置按鈕
    const resetBtn = document.getElementById("reset-btn");
    if (resetBtn) {
        resetBtn.addEventListener("click", () => {
            resetPageToInitialState();
            showToast("已清空重置當前頁面至初始狀態");
        });
    }

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

    // 下載 HTML 視覺化報表
    const downloadHtmlBtn = document.getElementById("download-html-btn");
    if (downloadHtmlBtn) {
        downloadHtmlBtn.addEventListener("click", () => {
            if (!currentJobId) {
                showToast("當前無可供下載的 HTML 視覺化報表，請先上傳檔案並執行審算", "error");
                return;
            }
            window.location.href = `/api/reports/download_html?job_id=${currentJobId}`;
        });
    }

    // 下載 Excel 報表
    document.getElementById("download-excel-btn").addEventListener("click", () => {
        if (!currentJobId) {
            showToast("當前無可供下載的審計報表，請先上傳檔案並執行審算", "error");
            return;
        }
        window.location.href = `/api/reports/download?job_id=${currentJobId}`;
    });

    // 啟動審核
    document.getElementById("run-audit-btn").addEventListener("click", startAudit);

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

    const uploadUrl = currentJobId ? `/api/upload?job_id=${currentJobId}` : "/api/upload";

    try {
        const res = await fetch(uploadUrl, { method: "POST", body: formData });
        const data = await res.json();

        if (data.status === "success" && data.job_id) {
            currentJobId = data.job_id;
            sessionStorage.setItem("currentJobId", currentJobId);
            document.getElementById("file-status-text").textContent = `已成功上傳 ${data.uploaded.length} 個檔案 (任務 ID: ${currentJobId})`;
            showToast(data.message || "上傳完成！點擊「啟動全自動 LLM 合規審核」即可進行比對。");
        }
    } catch (err) {
        showToast("上傳失敗: " + err.message, "error");
    }
}

async function startAudit() {
    if (!currentJobId) {
        showToast("請先上傳 XML 報文檔案再點擊啟動審核", "error");
        return;
    }

    const btn = document.getElementById("run-audit-btn");
    const progressBox = document.getElementById("progress-box");
    const progressFill = document.getElementById("progress-fill");
    const progressPercent = document.getElementById("progress-percent");

    btn.disabled = true;
    progressBox.classList.remove("hidden");

    try {
        const res = await fetch(`/api/audit/run?job_id=${currentJobId}`, { method: "POST" });
        const data = await res.json();

        if (data.status === "error") {
            showToast(data.message || "啟動失敗", "error");
            btn.disabled = false;
            progressBox.classList.add("hidden");
            return;
        }

        // 輪詢該任務專屬狀態
        const interval = setInterval(async () => {
            const statusRes = await fetch(`/api/audit/status?job_id=${currentJobId}`);
            const status = await statusRes.json();

            progressFill.style.width = `${status.progress}%`;
            progressPercent.textContent = `${status.progress}%`;

            if (!status.is_running) {
                clearInterval(interval);
                btn.disabled = false;
                progressBox.classList.add("hidden");
                showToast("合規審核作業完成！已更新最新數據。");
                await initApp();
            }
        }, 1500);

    } catch (err) {
        showToast("啟動審核失敗: " + err.message, "error");
        btn.disabled = false;
        progressBox.classList.add("hidden");
    }
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

function showToast(message, type = "info") {
    let toastContainer = document.getElementById("toast-container");
    if (!toastContainer) {
        toastContainer = document.createElement("div");
        toastContainer.id = "toast-container";
        toastContainer.style.cssText = `
            position: fixed;
            top: 24px;
            right: 24px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 10px;
            pointer-events: none;
        `;
        document.body.appendChild(toastContainer);
    }

    const toast = document.createElement("div");
    toast.style.cssText = `
        background: ${type === "error" ? "rgba(220, 38, 38, 0.92)" : "rgba(30, 41, 59, 0.92)"};
        color: #ffffff;
        padding: 12px 20px;
        border-radius: 8px;
        border: 1px solid ${type === "error" ? "rgba(248, 113, 113, 0.5)" : "rgba(96, 165, 250, 0.3)"};
        backdrop-filter: blur(8px);
        box-shadow: 0 10px 25px rgba(0, 0, 0, 0.5);
        font-size: 0.9rem;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 10px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        opacity: 0;
        transform: translateY(-12px);
        pointer-events: auto;
    `;
    
    const iconName = type === "error" ? "alert-circle" : "check-circle";
    toast.innerHTML = `<i data-lucide="${iconName}" style="width:18px;height:18px;flex-shrink:0;"></i><span>${escapeHtml(message)}</span>`;
    toastContainer.appendChild(toast);
    lucide.createIcons();

    setTimeout(() => {
        toast.style.opacity = "1";
        toast.style.transform = "translateY(0)";
    }, 10);

    setTimeout(() => {
        toast.style.opacity = "0";
        toast.style.transform = "translateY(-12px)";
        setTimeout(() => toast.remove(), 300);
    }, 3200);
}
