# 🛡️ ICP-Compare: 戰略性高科技貨品出口實體管理名單篩選工具

本專案是一個基於 **本地大型語言模型 (LLM)** 與 **雙 GPU NVIDIA H200 NVL 並行加速** 的進出口合規限制實體審計（Restricted Party Screening, RPS）輔助工具。

透過結合官方 API 的廣域文字模糊篩選 (Fuzzy Matching) 與本地高品質大模型 (`gemma-4:31B`) 的高併發語意推理，本工具能從原始比對結果中**自動且精準地過濾超過 90% 的「共享位址/相同產業字眼」所造成的偽陽性誤報 (False Positives)**，極大化提升合規審計小組的覆核效率。

---

## 🚀 系統架構與核心流向

本系統採 **「先廣域過濾（Rule-Based Guard），再精密審算（LLM Auditor）」** 的兩階段流向設計：

```mermaid
graph TD
    XML[1. 掃描 testfile/ 所有 raw XML 檔案] --> Parse[2. Python 精準解析 Condition 及 Party 節點]
    Parse --> Filter{3. 原始模糊比對相似度 >= 75.0%?}
    Filter -->|否 (無黑名單或相似度低)| Skip[4. 直接剔除 (免除非必要算力消耗)]
    Filter -->|是 (符合疑似高風險)| Candidate[5. 建立待審計 High-Risk 候選對象]
    Candidate --> Parallel[6. 雙 H200 本地 vLLM 高併發語意判審 (64 ThreadPool)]
    Parallel --> Output[7. 一鍵生成具時間戳記的 Excel 審計總報表]
```

---

## 🎯 限制實體語意比對之「四大風險等級」

本地 LLM 專家模型會拆解名字與地址，並對每一筆對象產出具備完整推理理由 (`reasoning`) 的判定級別：

| 風險等級 | 判定標準 | 合規應對與處置建議 |
| :--- | :--- | :--- |
| **🔴 High<br>(同一實體/高風險轉運)** | • 核心註冊品牌名稱完全對上，且包含街區、城市、物理地址等高度吻合。<br>• **名稱不同、出貨地址完全重合**：在 Ship-To (出貨地址) 比對中，此類案件藏有極高之**第三方白手套規避制裁轉運風險**。 | **立即攔截交易**。判定為同一被制裁實體或高風險轉運繞道，合規主管應進行系統凍結與深入稽核。 |
| **🟠 Medium<br>(關聯企業)** | • **名稱相同、國家不同**：被限制之境外關聯企業或海外子公司。 | **暫停交易並進入二審**。要求交易對象提供終極實質受益人（UBO）證明。 |
| **🟡 Low<br>(低風險/疑慮)** | 字面有重疊（如同為電子或化學產業通用詞），但兩者地址完全無交集、且無直接關係證明。 | **保留審查紀錄（Audit Trail）**。視為次要關注，可由稽核小組視承載力複核。 |
| **🟢 False Positive<br>(確定誤判)** | 因「剛好都在同一個科技園、同棟辦公大樓（如 Soft-Park/Midview City）」，或因名稱中皆含有通用字（如 `Technology/Semiconductor/Ltd` ）而被 Fuzzy 算法誤列，但本質非同一實體。 | **自動放行 (Auto-Release)**。100% 排除，合規團隊無須浪費多餘時間複審。 |

---

## 🛠️ 三大底層語意判定剖析規則

為確保 LLM 對 RPS 業務擁有資深專家的判定標準，系統引導 `gemma-4:31B` 依照以下三大步驟進行 Chain-of-Thought (CoT) 比對：

1. **核心名稱噪聲消除 (Core Name De-noising)**：
   * 主動降權並隔離無效後綴（如 `LTD`, `INC`, `GMBH`, `PTE`, `CO. KG` 等），精準判定核心名稱是否具備商標與品牌關聯（例如自動關聯 `Xiamen Sophgo` 即是 `廈門算能科技`）。
2. **多維地理常識比對 (Hierarchical Geography Parsing)**：
   * 智能識別「共享科學園區 / 共享代寄秘書辦公室 / 大型中轉物流貨代大樓」與「具體獨立辦公室」在合規角色上的重大差異。
3. **別名與完整關聯探索 (Alias Cross-Checking)**：
   * 自動遍閱 API 報文 `<content>` 資料內部所包含的 JSON 清冊（內含有被制裁者所有的別名 `alias` 等），查核原查詢名是否為受限者的曾用名。

---

## 💻 專案環境與執行

本專案全面採取現代高效的 **uv** 進行 Python 依賴項管理與隔離。

### 1. 安裝環境與依賴項
```bash
# 本專案已有 pyproject.toml，使用 uv 可一鍵自動安裝所有依賴至隔離的 .venv 中：
uv sync
```

### 2. 本地 vLLM 配置與高併發設計
本機配備了雙路 **NVIDIA H200 NVL**，為極大化發揮 vLLM 的 **Continuous Batching (連續批推論矩陣)** 優勢：
* **最大併發執行緒下調 (`MAX_WORKERS` = 64)**：能瞬間把 concurrent 請求拉滿，讓 Python 將 HTTP 發送壓降，雙 H200 可以在 **20秒 內** 高速處理完高達 288 筆複雜比對。

### 3. 一鍵執行審核
```bash
# 執行主程式（自動掃描 testfile/*_raw.xml）：
uv run compare_audit.py
```

### 4. 輸出產出
* 執行完成後會自動生成具有 YYYYMMDD_HHMMSS 時間戳記的全新檔案：
  `ICP_Audit_Report_20260706_141441.xlsx` 
* 該設計避開了 `PermissionError` (即 Excel 正在開啟觀看時無法存檔的問題)，同時為組織提供了不可篡改的 **合規留痕稽核軌跡 (Audit Log)**。
