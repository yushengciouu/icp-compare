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
                showAllPassState(data.stats);
                const statusEl = document.getElementById("file-status-text");
                if (statusEl) {
                    statusEl.innerHTML = `<span class="status-badge status-badge-pass"><i data-lucide="shield-check" style="width:14px;height:14px;display:inline;"></i> ✅ 審算完成：全數合格放行 (0 筆疑慮)</span>`;
                    lucide.createIcons();
                }
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
        file_name: "",
        last_updated: ""
    });

    document.getElementById("file-status-text").textContent = "等待上傳 XML 報文檔案";
    const mergedBtn = document.getElementById("download-excel-merged-btn");
    if (mergedBtn) {
        mergedBtn.innerHTML = `<i data-lucide="file-spreadsheet"></i> 匯出 Excel 報表`;
    }
    const splitBtn = document.getElementById("download-excel-split-btn");
    if (splitBtn) {
        splitBtn.style.display = "none";
    }
    showEmptyState("當前畫面已清空，請上傳 XML 報文檔案開始進行合規審算");
}

function updateKPIs(stats) {
    if (!stats) return;
    const highCount = Number(stats.high || 0);
    const mediumCount = Number(stats.medium || 0);
    const lowCount = Number(stats.low || 0);

    document.getElementById("kpi-total").textContent = stats.total || 0;
    document.getElementById("kpi-high").textContent = highCount;
    document.getElementById("kpi-medium").textContent = mediumCount;
    document.getElementById("kpi-low").textContent = lowCount;

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
    document.getElementById("count-low").textContent = lowCount;
    
    if (stats.file_name) {
        document.getElementById("file-status-text").textContent = `任務備忘: ${stats.file_name} (${stats.last_updated})`;
    }

    const reportCount = Number(stats.report_count || 1);
    const mergedBtn = document.getElementById("download-excel-merged-btn");
    const splitBtn = document.getElementById("download-excel-split-btn");
    if (mergedBtn) {
        if (reportCount > 1) {
            mergedBtn.innerHTML = `<i data-lucide="file-spreadsheet"></i> 匯出合併總 Excel`;
            if (splitBtn) {
                splitBtn.style.display = "inline-flex";
                splitBtn.innerHTML = `<i data-lucide="file-archive"></i> 匯出各檔案 Excel (${reportCount} 份 ZIP)`;
            }
        } else {
            mergedBtn.innerHTML = `<i data-lucide="file-spreadsheet"></i> 匯出 Excel 報表`;
            if (splitBtn) {
                splitBtn.style.display = "none";
            }
        }
        lucide.createIcons();
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
        default: return "badge-low";
    }
}

function getBadgeLabel(level) {
    switch (level) {
        case "High": return "🔴 High (同一實體/轉運風險)";
        case "Medium": return "🟠 Medium (關聯企業)";
        case "Low": return "🟡 Low (低風險/可放行)";
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
                <i data-lucide="brain-circuit"></i> 專家模型審計推理理由：
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
            ${level === 'Low' ? `
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

    // 下載合併總 Excel 報表
    const downloadMergedBtn = document.getElementById("download-excel-merged-btn");
    if (downloadMergedBtn) {
        downloadMergedBtn.addEventListener("click", () => {
            if (!currentJobId) {
                showToast("當前無可供下載的審計報表，請先上傳檔案並執行審算", "error");
                return;
            }
            window.location.href = `/api/reports/download?job_id=${currentJobId}&mode=merged`;
        });
    }

    // 下載個別檔案 Excel (ZIP 壓縮包)
    const downloadSplitBtn = document.getElementById("download-excel-split-btn");
    if (downloadSplitBtn) {
        downloadSplitBtn.addEventListener("click", () => {
            if (!currentJobId) {
                showToast("當前無可供下載的審計報表，請先上傳檔案並執行審算", "error");
                return;
            }
            window.location.href = `/api/reports/download?job_id=${currentJobId}&mode=individual`;
        });
    }

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

    // 每次上傳新檔案時，產生全新 job_id 進行完整隔離，避免受到上一輪歷史分析干擾
    const uploadUrl = "/api/upload";

    try {
        const res = await fetch(uploadUrl, { method: "POST", body: formData });
        const data = await res.json();

        if (data.status === "success" && data.job_id) {
            currentJobId = data.job_id;
            sessionStorage.setItem("currentJobId", currentJobId);
            
            const statusEl = document.getElementById("file-status-text");
            statusEl.innerHTML = `<span class="status-badge status-badge-ready"><i data-lucide="file-check" style="width:14px;height:14px;display:inline;"></i> 已上傳 ${data.uploaded.length} 個檔案 (就緒待審)</span>`;
            
            showEmptyState("檔案已上傳就緒！請點擊上方「啟動自動 LLM 合規審核」開始進行比對", "ready");
            showToast(`已成功上傳 ${data.uploaded.length} 個 XML 檔案！點擊「啟動自動 LLM 合規審核」即可進行比對。`, "info");
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
                await initApp();
                if (allRecords.length === 0) {
                    showToast("🎉 合規審算完成！本批次 XML 全數合格放行，無任何疑慮實體。", "success");
                } else {
                    showToast(`合規審查作業完成！共檢出 ${allRecords.length} 筆比對紀錄。`, "info");
                }
            }
        }, 1500);

    } catch (err) {
        showToast("啟動審核失敗: " + err.message, "error");
        btn.disabled = false;
        progressBox.classList.add("hidden");
    }
}

function showAllPassState(stats) {
    const container = document.getElementById("cards-container");
    const fileName = (stats && stats.file_name) ? stats.file_name : "XML 報文批次";
    const updateTime = (stats && stats.last_updated) ? stats.last_updated : new Date().toLocaleString();
    const reportCount = Number(stats && stats.report_count ? stats.report_count : 1);

    const actionButtonsHtml = reportCount > 1 ? `
        <button class="btn btn-primary" id="pass-export-merged-btn">
            <i data-lucide="file-spreadsheet"></i> 匯出合併總 Excel (全 ${reportCount} 檔整合)
        </button>
        <button class="btn btn-outline" id="pass-export-split-btn" style="background: rgba(59, 130, 246, 0.15); color: #93C5FD; border-color: rgba(59, 130, 246, 0.4);">
            <i data-lucide="file-archive"></i> 匯出各檔案獨立 Excel (共 ${reportCount} 份打包 .ZIP)
        </button>
        <button class="btn btn-outline" id="pass-export-html-btn">
            <i data-lucide="file-code"></i> 匯出 HTML 視覺化證明
        </button>
    ` : `
        <button class="btn btn-primary" id="pass-export-merged-btn">
            <i data-lucide="file-spreadsheet"></i> 匯出合格存查 Excel 報表
        </button>
        <button class="btn btn-outline" id="pass-export-html-btn">
            <i data-lucide="file-code"></i> 匯出 HTML 視覺化證明
        </button>
    `;

    container.innerHTML = `
        <div class="all-pass-card glass-card">
            <div class="pass-icon-glow">
                <i data-lucide="shield-check"></i>
            </div>
            <div class="pass-content">
                <div class="pass-badge">
                    <i data-lucide="check-circle-2" style="width: 14px; height: 14px;"></i>
                    合規判定：100% 全數合格通過 (Pass)
                </div>
                <h2>🎉 合規審核完成：未命中任何限制實體！</h2>
                <p class="pass-desc">
                    本批次所上傳之 XML 報文檔案，經廣域模糊篩選比對結果<strong>全數低於 75% 警戒門檻</strong>，無任何高風險 (High) 或中風險 (Medium) 限制黑名單疑慮實體，<strong>可安心直接放行</strong>。
                </p>
                <div class="pass-meta-grid">
                    <div class="meta-item">
                        <span class="meta-label"><i data-lucide="file-text" style="width:13px;height:13px;display:inline;"></i> 任務報表</span>
                        <span class="meta-val">${escapeHtml(fileName)}</span>
                    </div>
                    <div class="meta-item">
                        <span class="meta-label"><i data-lucide="clock" style="width:13px;height:13px;display:inline;"></i> 審算完成時間</span>
                        <span class="meta-val">${escapeHtml(updateTime)}</span>
                    </div>
                    <div class="meta-item">
                        <span class="meta-label"><i data-lucide="file-check" style="width:13px;height:13px;display:inline;"></i> 存查憑證狀態</span>
                        <span class="meta-val text-green">✅ 已產出合併總表與 ${reportCount} 份個別合格 Excel 報表</span>
                    </div>
                </div>
                <div class="pass-actions">
                    ${actionButtonsHtml}
                </div>
            </div>
        </div>
    `;

    const dlMerged = document.getElementById("pass-export-merged-btn");
    if (dlMerged) {
        dlMerged.addEventListener("click", () => {
            const btn = document.getElementById("download-excel-merged-btn");
            if (btn) btn.click();
        });
    }
    const dlSplit = document.getElementById("pass-export-split-btn");
    if (dlSplit) {
        dlSplit.addEventListener("click", () => {
            const btn = document.getElementById("download-excel-split-btn");
            if (btn) btn.click();
        });
    }
    const dlHtml = document.getElementById("pass-export-html-btn");
    if (dlHtml) {
        dlHtml.addEventListener("click", () => {
            const btn = document.getElementById("download-html-btn");
            if (btn) btn.click();
        });
    }

    lucide.createIcons();
}

function showEmptyState(msg, type = "info") {
    const container = document.getElementById("cards-container");
    const iconName = type === "ready" ? "arrow-up-circle" : "inbox";
    const iconColor = type === "ready" ? "var(--color-accent-blue)" : "var(--text-muted)";
    container.innerHTML = `
        <div class="loading-state">
            <i data-lucide="${iconName}" style="width: 52px; height: 52px; color: ${iconColor}; opacity: 0.85;"></i>
            <p style="margin-top: 14px; font-size: 0.95rem; color: var(--text-main); font-weight: 500;">${escapeHtml(msg)}</p>
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
    const isSuccess = type === "success";
    const isError = type === "error";

    const bgColor = isSuccess ? "rgba(6, 78, 59, 0.95)" : isError ? "rgba(220, 38, 38, 0.92)" : "rgba(30, 41, 59, 0.92)";
    const borderColor = isSuccess ? "rgba(52, 211, 153, 0.7)" : isError ? "rgba(248, 113, 113, 0.5)" : "rgba(96, 165, 250, 0.3)";
    const glowColor = isSuccess ? "0 10px 30px rgba(16, 185, 129, 0.4)" : "0 10px 25px rgba(0, 0, 0, 0.5)";

    toast.style.cssText = `
        background: ${bgColor};
        color: #ffffff;
        padding: 14px 22px;
        border-radius: 10px;
        border: 1px solid ${borderColor};
        backdrop-filter: blur(12px);
        box-shadow: ${glowColor};
        font-size: 0.92rem;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 12px;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        opacity: 0;
        transform: translateY(-12px);
        pointer-events: auto;
    `;
    
    const iconName = isSuccess ? "shield-check" : isError ? "alert-circle" : "info";
    toast.innerHTML = `<i data-lucide="${iconName}" style="width:20px;height:20px;flex-shrink:0;${isSuccess ? 'color:#34D399;' : ''}"></i><span>${escapeHtml(message)}</span>`;
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
    }, 4200);
}
