// ==UserScript==
// @name         深圳大学平时成绩&期末成绩查询
// @namespace    http://tampermonkey.net/
// @version      4.19
// @description  开发者模式监听页面真实请求，支持页面内表格展示
// @author       流年.
// @match        https://ehall.szu.edu.cn/jwapp/sys/cjcx/*
// @match        https://ehall-443.webvpn.szu.edu.cn/jwapp/sys/cjcx/*
// @connect      ehall.szu.edu.cn
// @connect      ehall-443.webvpn.szu.edu.cn
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @license      MIT
// ==/UserScript==

(function() {
    'use strict';

    let scriptState = {
        isRunning: false,
        courseData: [],
        container: null,
        studentId: null,
        studentName: null,
        devMode: false,
        isProbing: false,
        queryProgress: {
            active: false,
            percent: 0,
            message: '准备就绪',
            detail: '',
            updatedAt: null
        },
        rawData: {
            initialCourses: null,
            queryResults: [],  // 存储轮询结果
            probeResults: null,
            networkCaptures: []
        },
        networkMonitor: {
            installed: false,
            active: false,
            originalFetch: null,
            originalXHROpen: null,
            originalXHRSend: null
        },
        inlineScoreTab: {
            installed: false,
            tab: null,
            panel: null
        },
        tableSort: {
            field: 'courseName',
            direction: 'asc'
        }
    };

    // [优化] 注入优化的核心样式
    GM_addStyle(`
        /* Main container and general layout */
        #score-query-container {
            position: fixed;
            top: 20px;
            right: 20px;
            width: 500px;
            max-width: calc(100vw - 40px);
            background: #f9f9f9;
            border-radius: 16px;
            padding: 20px;
            z-index: 99999;
            box-shadow: 0 8px 24px rgba(0,0,0,0.15);
            transition: transform 0.4s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.4s;
            display: flex;
            flex-direction: column;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
        }
        #score-query-container.hidden {
            transform: translateX(110%);
            opacity: 0;
            pointer-events: none;
        }

        /* Header */
        .sq-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 1px solid #e0e0e0;
        }
        .sq-header h3 {
            margin: 0;
            font-size: 1.1rem;
            font-weight: 600;
            color: #212121;
        }
        .sq-close-btn {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 32px;
            height: 32px;
            border: none;
            background: #e0e0e0;
            border-radius: 50%;
            cursor: pointer;
            transition: background-color 0.2s, transform 0.2s;
        }
        .sq-close-btn:hover {
            background-color: #d1d1d1;
            transform: rotate(90deg);
        }
        .sq-close-btn svg {
            width: 14px;
            height: 14px;
            stroke: #555;
        }

        /* Main content area */
        .sq-content {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
        }

        /* Action Buttons */
        .sq-actions {
            display: flex;
            gap: 12px;
            margin-bottom: 16px;
        }
        .sq-btn {
            flex-grow: 1;
            padding: 12px;
            border: none;
            border-radius: 8px;
            color: #fff;
            font-size: 0.9rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.25s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }
        .sq-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }
        .sq-btn:disabled {
            background: #bdbdbd !important;
            cursor: not-allowed;
            box-shadow: none;
            transform: none;
        }
        #start-query {
            background: linear-gradient(135deg, #43A047 0%, #66BB6A 100%);
        }
        #export-scores {
            background: linear-gradient(135deg, #1E88E5 0%, #42A5F5 100%);
        }

        /* Progress and Status */
        .progress-container {
            margin-bottom: 8px;
            display: none;
        }
        .progress-container.active {
            display: block;
        }
        .progress-container.completed {
            display: none;
        }
        .progress-bar {
            height: 6px;
            background: #e0e0e0;
            border-radius: 3px;
            overflow: hidden;
        }
        .progress {
            height: 100%;
            background: linear-gradient(90deg, #43A047, #81C784);
            width: 0%;
            transition: width 0.3s ease-in-out;
        }
        #status {
            margin-bottom: 8px;
            font-size: 0.85rem;
            color: #616161;
            text-align: center;
            min-height: 20px;
        }

        /* Results Area */
        #score-results {
            max-height: 400px;
            overflow: auto;
            margin: 0 -12px;
            padding: 4px 12px;
        }
        .score-summary-card {
            background: #e3f2fd;
            padding: 16px;
            border-radius: 8px;
            margin-bottom: 16px;
            border: 1px solid #bbdefb;
        }
        .score-summary-title {
            font-size: 1.2rem;
            font-weight: 700;
            color: #1565c0;
            margin-bottom: 12px;
        }
        .score-summary-grid {
            display: grid;
            grid-template-columns: minmax(220px, 1fr) minmax(220px, 1fr);
            gap: 14px;
            align-items: stretch;
            margin-bottom: 14px;
        }
        .score-summary-panel {
            width: 100%;
            height: 100%;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            background: #fff;
            border: 1px solid #d6e3ef;
            border-radius: 8px;
            overflow: hidden;
        }
        .score-summary-panel-title {
            padding: 9px 12px;
            background: #f5fbff;
            border-bottom: 1px solid #d6e3ef;
            color: #455a64;
            font-size: 0.86rem;
            font-weight: 700;
        }
        .score-summary-table {
            width: 100%;
            flex: 1;
            border-collapse: collapse;
            table-layout: fixed;
            background: #fff;
            font-size: 0.84rem;
        }
        .score-summary-table th,
        .score-summary-table td {
            padding: 8px 12px;
            border-top: 1px solid #eef3f7;
            text-align: left;
            line-height: 1.4;
        }
        .score-summary-table tr:first-child th,
        .score-summary-table tr:first-child td {
            border-top: none;
        }
        .score-summary-table th {
            color: #607d8b;
            font-weight: 600;
        }
        .score-summary-table td {
            color: #263238;
            font-weight: 700;
            text-align: right;
            font-variant-numeric: tabular-nums;
        }
        .score-chart-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(280px, 1fr));
            gap: 14px;
        }
        .score-chart-card {
            width: 100%;
            box-sizing: border-box;
            min-width: 0;
        }
        .score-chart-card svg {
            width: 100%;
            height: auto;
            display: block;
        }
        #score-query-container .score-chart-grid {
            grid-template-columns: 1fr;
            gap: 10px;
        }
        .score-semester-section {
            margin-bottom: 18px;
        }
        .score-semester-header {
            margin: 12px 0 8px 0;
            padding: 8px 0 4px 0;
            border-bottom: 2px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: sticky;
            top: -4px;
            background: #f9f9f9;
            z-index: 10;
        }
        .score-semester-header h4 {
            margin: 0;
            color: #333;
            font-size: 0.95rem;
        }
        .score-semester-header span {
            font-weight: 700;
            color: #4caf50;
            font-size: 0.85rem;
        }
        .score-table-wrap {
            width: 100%;
            overflow-x: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            background: #fff;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
        }
        .score-table {
            width: 100%;
            min-width: 1180px;
            border-collapse: collapse;
            background: #fff;
            font-size: 0.78rem;
            table-layout: fixed;
        }
        .score-table col.score-col-course { width: auto; }
        .score-table col.score-col-nature { width: 104px; }
        .score-table col.score-col-credit { width: 78px; }
        .score-table col.score-col-total { width: 86px; }
        .score-table col.score-col-grade { width: 96px; }
        .score-table col.score-col-regular { width: 104px; }
        .score-table col.score-col-final { width: 104px; }
        .score-table col.score-col-regular-coeff { width: 156px; }
        .score-table col.score-col-final-coeff { width: 156px; }
        @media (max-width: 900px) {
            .score-summary-grid,
            .score-chart-grid {
                grid-template-columns: 1fr;
            }
        }
        .score-table thead tr {
            background: #3498db;
            color: #fff;
        }
        .score-table th,
        .score-table td {
            padding: 9px 10px;
            text-align: center;
            border-top: 1px solid #e6e6e6;
            vertical-align: middle;
            line-height: 1.45;
            white-space: nowrap;
        }
        .score-table th {
            border-top: none;
            font-weight: 700;
            white-space: nowrap;
        }
        .score-table th.sortable {
            cursor: pointer;
            user-select: none;
            transition: background-color 0.18s ease;
        }
        .score-table th.sortable:hover {
            background: #2d8dcc;
        }
        .score-table th.active-sort {
            background: #2384c4;
        }
        .score-table .sort-indicator {
            display: inline-block;
            width: 1em;
            margin-left: 4px;
            font-size: 0.75rem;
            line-height: 1;
            vertical-align: middle;
        }
        .score-table th:first-child,
        .score-table td:first-child {
            text-align: left;
        }
        .score-table tbody tr {
            transition: background-color 0.2s ease;
        }
        .score-table tbody tr:hover {
            background: #f5f9ff;
        }
        .score-table .course-name-cell {
            font-weight: 600;
            color: #263238;
            white-space: normal;
            word-break: break-word;
            overflow-wrap: anywhere;
        }
        .score-table .score-number {
            font-weight: 700;
            color: #d81b60;
        }
        .score-table .score-muted {
            color: #8a8a8a;
        }
        .score-table .score-coeff {
            white-space: nowrap;
            color: #455a64;
        }
        .szu-inline-score-panel {
            padding: 16px;
            background: #f9f9f9;
            min-height: 240px;
            box-sizing: border-box;
        }
        .szu-inline-score-toolbar {
            display: flex;
            gap: 10px;
            align-items: center;
            margin-bottom: 14px;
        }
        .szu-inline-score-toolbar button {
            padding: 8px 18px;
            font-size: 14px;
            background: #e6a23c;
            color: #fff;
            border: none;
            border-radius: 20px;
            cursor: pointer;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
        }
        .szu-inline-score-toolbar button:hover {
            background: #cf8f27;
        }
        .szu-inline-score-hint {
            color: #666;
            font-size: 13px;
        }
        .szu-inline-progress-card {
            margin: 0 0 14px;
            padding: 12px 14px;
            background: #fff;
            border: 1px solid #e3e8ef;
            border-radius: 8px;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.04);
        }
        .szu-inline-progress-head {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            margin-bottom: 8px;
            color: #455a64;
            font-size: 14px;
            font-weight: 600;
        }
        .szu-inline-progress-percent {
            flex: 0 0 auto;
            color: #009688;
            font-weight: 700;
            font-variant-numeric: tabular-nums;
        }
        .szu-inline-progress-track {
            height: 8px;
            overflow: hidden;
            background: #e8eef3;
            border-radius: 999px;
        }
        .szu-inline-progress-fill {
            height: 100%;
            width: 0%;
            background: linear-gradient(90deg, #26a69a, #66bb6a);
            border-radius: inherit;
            transition: width 0.9s cubic-bezier(0.22, 1, 0.36, 1);
            will-change: width;
        }
        .szu-inline-progress-detail {
            margin-top: 8px;
            color: #78909c;
            font-size: 12px;
            line-height: 1.4;
        }
        .szu-inline-score-empty {
            text-align: center;
            padding: 28px 16px;
            color: #777;
            font-size: 15px;
            background: #fff;
            border: 1px dashed #d0d0d0;
            border-radius: 8px;
        }
        .course-item {
            padding: 16px;
            background: #fff;
            border: 1px solid #e8e8e8;
            border-radius: 8px;
            margin-bottom: 12px;
            transition: box-shadow 0.2s, transform 0.2s;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px 16px;
        }
        .course-item:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        .course-item:last-child {
            margin-bottom: 0;
        }
        .course-header {
            grid-column: 1 / -1;
            margin-bottom: 8px;
            padding-bottom: 8px;
            border-bottom: 1px dashed #eee;
        }
        .course-header strong {
            font-size: 1.05rem;
            color: #333;
            display: block;
        }
        .course-header span {
            font-size: 0.8rem;
            color: #757575;
        }
        .course-detail {
            font-size: 0.85rem;
            color: #616161;
            line-height: 1.6;
        }
        .course-detail.full-width {
            grid-column: 1 / -1;
        }
        .score-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .final-score {
            font-weight: bold;
            color: #d81b60;
            font-size: 1rem;
        }
        .tag {
            display: inline-block;
            padding: 2px 6px;
            background: #f5f5f5;
            border-radius: 4px;
            font-size: 0.75rem;
            color: #666;
            margin-right: 4px;
        }
        #score-results::-webkit-scrollbar { width: 6px; }
        #score-results::-webkit-scrollbar-thumb { background: #bdbdbd; border-radius: 3px; }
        #score-results::-webkit-scrollbar-track { background: transparent; }

        /* Footer */
        .sq-footer {
            margin-top: 20px;
            padding-top: 12px;
            border-top: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.8rem;
            color: #757575;
        }
        .github-link {
            display: flex;
            align-items: center;
            gap: 6px;
            color: #757575;
            text-decoration: none;
            transition: color 0.2s;
        }
        .github-link:hover {
            color: #212121;
        }
        .github-link svg {
            width: 18px;
            height: 18px;
            fill: currentColor;
        }

        /* Toggle Button */
        #toggle-btn {
            position: fixed;
            top: 20px;
            right: 20px;
            width: 56px;
            height: 56px;
            background: linear-gradient(135deg, #43A047 0%, #66BB6A 100%);
            color: #fff;
            border: none;
            border-radius: 50%;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            z-index: 99998;
            box-shadow: 0 6px 18px rgba(67, 160, 71, 0.3);
            transition: all 0.25s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
            line-height: 1.2;
        }
        #toggle-btn:hover {
            box-shadow: 0 8px 24px rgba(67, 160, 71, 0.4);
            transform: translateY(-2px) scale(1.05);
        }

        /* Dev Mode Styles */
        .sq-dev-toggle {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 12px;
            padding: 8px 12px;
            background: #fff3e0;
            border-radius: 6px;
            font-size: 0.8rem;
            color: #e65100;
        }
        .sq-dev-toggle input[type="checkbox"] {
            cursor: pointer;
        }
        .sq-dev-toggle label {
            cursor: pointer;
            user-select: none;
        }
        .sq-dev-badge {
            background: #ff6d00;
            color: #fff;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 0.7rem;
            font-weight: 600;
        }
        #dev-raw-data {
            display: none;
            margin-top: 12px;
        }
        #dev-raw-data.visible {
            display: block;
        }
        .dev-query-list {
            max-height: 300px;
            overflow-y: auto;
        }
        .dev-query-item {
            margin-bottom: 8px;
            border: 1px solid #424242;
            border-radius: 4px;
            overflow: hidden;
        }
        .dev-query-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 6px 10px;
            background: #37474f;
            color: #fff;
            font-size: 0.8rem;
            cursor: pointer;
        }
        .dev-query-header:hover {
            background: #455a64;
        }
        .dev-query-badge {
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 0.7rem;
            font-weight: 600;
        }
        .dev-query-badge.pscj {
            background: #4CAF50;
        }
        .dev-query-badge.qmcj {
            background: #FF5722;
        }
        .dev-query-badge.count {
            background: #2196F3;
            margin-left: 6px;
        }
        .dev-query-body {
            display: none;
            background: #263238;
            color: #80cbc4;
            padding: 8px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 0.7rem;
            white-space: pre-wrap;
            word-break: break-all;
            max-height: 150px;
            overflow-y: auto;
        }
        .dev-query-body.expanded {
            display: block;
        }
        .dev-clear-btn {
            margin-top: 6px;
            padding: 4px 10px;
            background: #f44336;
            color: #fff;
            border: none;
            border-radius: 4px;
            font-size: 0.75rem;
            cursor: pointer;
            transition: background 0.2s;
        }
        .dev-clear-btn:hover {
            background: #d32f2f;
        }
        .dev-probe-actions {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 8px;
            margin-bottom: 8px;
        }
        .dev-probe-btn,
        .dev-download-btn {
            padding: 4px 10px;
            color: #fff;
            border: none;
            border-radius: 4px;
            font-size: 0.75rem;
            cursor: pointer;
            transition: background 0.2s, opacity 0.2s;
        }
        .dev-probe-btn {
            background: #3949ab;
        }
        .dev-probe-btn:hover {
            background: #283593;
        }
        .dev-download-btn {
            background: #1976d2;
        }
        .dev-download-btn:hover {
            background: #0d47a1;
        }
        .dev-probe-btn:disabled,
        .dev-download-btn:disabled {
            opacity: 0.55;
            cursor: not-allowed;
        }
        .dev-probe-status {
            padding: 8px 10px;
            margin-bottom: 8px;
            background: #fff8e1;
            border-left: 4px solid #f9a825;
            border-radius: 4px;
            color: #795548;
            font-size: 0.78rem;
            line-height: 1.45;
        }
        .dev-monitor-status {
            padding: 8px 10px;
            margin-bottom: 8px;
            background: #e3f2fd;
            border-left: 4px solid #1976d2;
            border-radius: 4px;
            color: #0d47a1;
            font-size: 0.78rem;
            line-height: 1.45;
        }
        .dev-data-section {
            margin-bottom: 12px;
        }
        .dev-data-section summary {
            cursor: pointer;
            padding: 8px 12px;
            background: #424242;
            color: #fff;
            border-radius: 6px;
            font-size: 0.85rem;
            font-weight: 500;
            user-select: none;
        }
        .dev-data-section summary:hover {
            background: #616161;
        }
        .dev-data-content {
            max-height: 200px;
            overflow-y: auto;
            background: #263238;
            color: #80cbc4;
            padding: 12px;
            border-radius: 0 0 6px 6px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 0.75rem;
            white-space: pre-wrap;
            word-break: break-all;
        }
        .dev-copy-btn {
            margin-top: 6px;
            padding: 4px 10px;
            background: #00897b;
            color: #fff;
            border: none;
            border-radius: 4px;
            font-size: 0.75rem;
            cursor: pointer;
            transition: background 0.2s;
        }
        .dev-copy-btn:hover {
            background: #00695c;
        }
    `);

    const toggleBtn = document.createElement('button');
    toggleBtn.id = 'toggle-btn';
    toggleBtn.innerHTML = '深大<br>成绩';
    document.body.appendChild(toggleBtn);

    function getStudentInfoFromPage() {
        const allTds = document.querySelectorAll('td');
        for (const td of allTds) {
            const text = td.textContent.trim();
            if (text === '学号' && td.nextElementSibling) {
                scriptState.studentId = td.nextElementSibling.textContent.trim();
            }
            if (text === '姓名' && td.nextElementSibling) {
                scriptState.studentName = td.nextElementSibling.textContent.trim();
            }
            if (scriptState.studentId && scriptState.studentName) {
                break;
            }
        }
    }

    function initContainer() {
        const container = document.createElement('div');
        container.id = 'score-query-container';
        container.className = 'hidden';
        container.innerHTML = `
            <div class="sq-header">
                <h3>深圳大学成绩查询助手</h3>
                <button class="sq-close-btn" title="关闭">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                </button>
            </div>

            <div class="sq-content">
                <div class="sq-dev-toggle" id="dev-toggle-container" style="display: none;">
                    <input type="checkbox" id="dev-mode-checkbox">
                    <label for="dev-mode-checkbox">开发者模式</label>
                    <span class="sq-dev-badge">DEV</span>
                </div>
                <div class="sq-actions">
                    <button id="start-query" class="sq-btn">开始查询</button>
                    <button id="export-scores" class="sq-btn" disabled>导出Excel</button>
                </div>
                <div class="progress-container">
                    <div id="status">准备就绪</div>
                    <div class="progress-bar"><div class="progress" id="progress"></div></div>
                </div>
                <div id="score-results"></div>
                <div id="dev-raw-data">
                    <details class="dev-data-section">
                        <summary>📋 初始课程列表数据</summary>
                        <div class="dev-data-content" id="dev-initial-data">暂无数据</div>
                        <button class="dev-copy-btn" data-target="dev-initial-data">复制到剪贴板</button>
                    </details>
                    <details class="dev-data-section">
                        <summary>🔄 轮询查询结果 (<span id="dev-query-count">0</span>条)</summary>
                        <div class="dev-query-list" id="dev-query-list">
                            <div style="padding:12px;color:#999;text-align:center;">暂无查询记录</div>
                        </div>
                        <button class="dev-copy-btn" id="dev-copy-all-queries">复制全部查询结果</button>
                        <button class="dev-clear-btn" id="dev-clear-queries">清空记录</button>
                    </details>
                    <details class="dev-data-section">
                        <summary>🧪 系数接口主动探测</summary>
                        <div class="dev-probe-actions">
                            <button class="dev-probe-btn" id="dev-run-probe">开始探测系数接口</button>
                            <button class="dev-copy-btn" data-target="dev-probe-data">复制探测结果</button>
                            <button class="dev-download-btn" id="dev-download-probe-results" disabled>下载探测结果</button>
                        </div>
                        <div class="dev-probe-status" id="dev-probe-status">尚未探测。请先确认已登录成绩查询页面，再点击开始探测。</div>
                        <div class="dev-data-content" id="dev-probe-data">暂无数据</div>
                    </details>
                    <details class="dev-data-section">
                        <summary>📡 页面请求监听 (<span id="dev-network-count">0</span>条)</summary>
                        <div class="dev-probe-actions">
                            <button class="dev-probe-btn" id="dev-start-network-monitor">开始监听</button>
                            <button class="dev-clear-btn" id="dev-stop-network-monitor" disabled>停止监听</button>
                            <button class="dev-copy-btn" data-target="dev-network-data">复制监听结果</button>
                            <button class="dev-download-btn" id="dev-download-network-captures" disabled>下载监听结果</button>
                            <button class="dev-clear-btn" id="dev-clear-network-captures">清空监听记录</button>
                        </div>
                        <div class="dev-monitor-status" id="dev-network-status">尚未监听。点击开始监听后，请在官方成绩页面点击“详情”等操作。</div>
                        <div class="dev-data-content" id="dev-network-data">暂无数据</div>
                    </details>
                </div>
            </div>

            <div class="sq-footer">
                <span>&copy; 2025 流年</span>
                <a href="https://github.com/Liunian2000/GradeInquiry4SZU/" target="_blank" class="github-link" title="查看源码">
                    <svg role="img" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><title>GitHub</title><path d="M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12"/></svg>
                    <span>GitHub</span>
                </a>
            </div>
        `;
        document.body.appendChild(container);
        scriptState.container = container;

        const startBtn = container.querySelector('#start-query');
        const exportBtn = container.querySelector('#export-scores');
        const statusEl = container.querySelector('#status');
        const progressEl = container.querySelector('#progress');
        const resultsEl = container.querySelector('#score-results');
        const closeBtn = container.querySelector('.sq-close-btn');
        const devToggleContainer = container.querySelector('#dev-toggle-container');
        const devModeCheckbox = container.querySelector('#dev-mode-checkbox');
        const devRawDataEl = container.querySelector('#dev-raw-data');

        closeBtn.addEventListener('click', () => container.classList.add('hidden'));

        // 开发者模式切换
        devModeCheckbox.addEventListener('change', (e) => {
            scriptState.devMode = e.target.checked;
            if (scriptState.devMode) {
                devRawDataEl.classList.add('visible');
                updateDevDataDisplay();
            } else {
                devRawDataEl.classList.remove('visible');
            }
        });

        // 复制按钮事件
        container.querySelectorAll('.dev-copy-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const targetId = btn.getAttribute('data-target');
                const targetEl = container.querySelector(`#${targetId}`);
                if (targetEl) {
                    const text = targetEl.textContent;
                    navigator.clipboard.writeText(text).then(() => {
                        const originalText = btn.textContent;
                        btn.textContent = '已复制!';
                        btn.style.background = '#4CAF50';
                        setTimeout(() => {
                            btn.textContent = originalText;
                            btn.style.background = '';
                        }, 1500);
                    }).catch(err => {
                        console.error('复制失败:', err);
                        alert('复制失败，请手动复制');
                    });
                }
            });
        });

        // 复制全部查询结果按钮
        container.querySelector('#dev-copy-all-queries').addEventListener('click', () => {
            const text = JSON.stringify(scriptState.rawData.queryResults, null, 2);
            navigator.clipboard.writeText(text).then(() => {
                const btn = container.querySelector('#dev-copy-all-queries');
                const originalText = btn.textContent;
                btn.textContent = '已复制!';
                btn.style.background = '#4CAF50';
                setTimeout(() => {
                    btn.textContent = originalText;
                    btn.style.background = '';
                }, 1500);
            }).catch(err => {
                console.error('复制失败:', err);
                alert('复制失败，请手动复制');
            });
        });

        // 清空查询记录按钮
        container.querySelector('#dev-clear-queries').addEventListener('click', () => {
            scriptState.rawData.queryResults = [];
            updateDevQueryDisplay();
        });

        // 系数接口主动探测
        container.querySelector('#dev-run-probe').addEventListener('click', async () => {
            if (scriptState.isProbing) return;

            const runProbeBtn = container.querySelector('#dev-run-probe');
            const probeStatusEl = container.querySelector('#dev-probe-status');
            const downloadBtn = container.querySelector('#dev-download-probe-results');

            scriptState.isProbing = true;
            runProbeBtn.disabled = true;
            downloadBtn.disabled = true;
            scriptState.rawData.probeResults = null;
            updateDevProbeDisplay();

            const updateProbeStatus = (message) => {
                if (probeStatusEl) probeStatusEl.textContent = message;
            };

            const updateProbeResults = (results) => {
                scriptState.rawData.probeResults = results;
                updateDevProbeDisplay();
                downloadBtn.disabled = !scriptState.rawData.probeResults;
            };

            try {
                const results = await runCoefficientEndpointProbe(updateProbeStatus, updateProbeResults);
                scriptState.rawData.probeResults = results;
                const officialHitCount = results.officialCoefficientProbe?.coefficientHits?.length || 0;
                updateProbeStatus(`探测完成：官方系数命中 ${officialHitCount} 条，扫描资源 ${results.discovery.resourcesScanned.length} 个，候选接口 ${results.candidates.length} 个，请下载 JSON 发回分析。`);
                updateDevProbeDisplay();
            } catch (err) {
                console.error('[深大成绩查询] 接口探测失败:', err);
                updateProbeStatus(`探测失败：${err.message}`);
                scriptState.rawData.probeResults = {
                    ...(scriptState.rawData.probeResults || {}),
                    state: 'failed',
                    failedAt: new Date().toISOString(),
                    error: err.message,
                    stack: err.stack
                };
                updateDevProbeDisplay();
            } finally {
                scriptState.isProbing = false;
                runProbeBtn.disabled = false;
                downloadBtn.disabled = !scriptState.rawData.probeResults;
            }
        });

        container.querySelector('#dev-download-probe-results').addEventListener('click', () => {
            if (!scriptState.rawData.probeResults) {
                alert('暂无探测结果，请先执行探测。');
                return;
            }
            downloadJsonFile(scriptState.rawData.probeResults, `szu-score-probe-${formatDateTimeForFilename(new Date())}.json`);
        });

        container.querySelector('#dev-start-network-monitor').addEventListener('click', () => {
            startNetworkMonitor();
        });

        container.querySelector('#dev-stop-network-monitor').addEventListener('click', () => {
            stopNetworkMonitor();
        });

        container.querySelector('#dev-download-network-captures').addEventListener('click', () => {
            if (!scriptState.rawData.networkCaptures.length) {
                alert('暂无监听结果，请先开始监听并操作官方页面。');
                return;
            }
            const payload = buildNetworkCaptureExport();
            downloadJsonFile(payload, `szu-score-network-captures-${formatDateTimeForFilename(new Date())}.json`);
        });

        container.querySelector('#dev-clear-network-captures').addEventListener('click', () => {
            scriptState.rawData.networkCaptures = [];
            updateDevNetworkDisplay();
        });

        startBtn.addEventListener('click', async () => {
            if (scriptState.isRunning) return;

            getStudentInfoFromPage();

            scriptState.isRunning = true;
            startBtn.disabled = true;
            exportBtn.disabled = true;
            scriptState.courseData = [];
            resultsEl.innerHTML = '';
            renderInlineScorePanel();
            progressEl.style.width = '0%';
            // 显示进度条区域
            const progressContainer = container.querySelector('.progress-container');
            progressContainer.classList.remove('completed');
            progressContainer.classList.add('active');
            setQueryProgress(2, '正在获取课程列表...', '正在连接成绩查询接口。');

            try {
                // 1. 获取初始课程列表
                const initialCourses = await fetchInitialCourseList();
                if (!initialCourses || initialCourses.length === 0) {
                    setQueryProgress(100, '未找到任何课程记录，请确认当前学期有成绩。', '', false);
                    return;
                }

                // 2. 优先获取课程系数（新增逻辑）
                setQueryProgress(8, `正在获取课程系数 (0/${initialCourses.length})...`, '正在读取每门课的成绩项配置。');
                const coefficientMap = new Map();
                
                // 分批并发获取系数
                const batchSize = 5;
                for (let i = 0; i < initialCourses.length; i += batchSize) {
                    const batch = initialCourses.slice(i, i + batchSize);
                    await Promise.all(batch.map(async (course) => {
                        if (course.JXBID) {
                            const coeffs = await fetchCourseCoefficients(course.JXBID);
                            if (coeffs) {
                                coefficientMap.set(course.JXBID, coeffs);
                            }
                        }
                    }));
                    const coefficientProgress = 8 + Math.round((Math.min(i + batchSize, initialCourses.length) / initialCourses.length) * 12);
                    setQueryProgress(coefficientProgress, `正在获取课程系数 (${Math.min(i + batchSize, initialCourses.length)}/${initialCourses.length})...`, '正在读取每门课的成绩项配置。');
                    await new Promise(r => setTimeout(r, 50)); // 稍微延时防止请求过快
                }

                // 3. 初始化课程Map，并根据系数判断需要查询哪些成绩
                const courseMap = new Map();
                let needPscjCount = 0;  // 需要查询平时成绩的课程数
                let needQmcjCount = 0;  // 需要查询期末成绩的课程数
                
                initialCourses.forEach(course => {
                    const key = course.KCM + course.XNXQDM_DISPLAY;
                    
                    // 初始化成绩
                    course.PSCJ = 'N/A';
                    course.QMCJ = 'N/A';
                    
                    // 检查是否获取到了官方系数
                    const officialCoeffs = coefficientMap.get(course.JXBID);
                    
                    if (officialCoeffs) {
                        // 使用官方系数
                        course.PSCJXS = officialCoeffs.pscjxs;
                        course.QMCJXS = officialCoeffs.qmcjxs;
                        course._pscjxsNum = parseFloat(officialCoeffs.pscjxs);
                        course._qmcjxsNum = parseFloat(officialCoeffs.qmcjxs);
                        course._coefficientsSource = 'official'; // 标记来源
                        course._coefficientsInferred = false;
                        
                        // 根据系数优化查询需求
                        // 如果系数为0，则不需要查询对应成绩
                        course._needPscj = course._pscjxsNum > 0;
                        course._needQmcj = course._qmcjxsNum > 0;
                        
                        if (!course._needPscj) course.PSCJ = '-';
                        if (!course._needQmcj) course.QMCJ = '-';
                        
                        console.log(`[系数获取] ${course.KCM}: 使用接口系数 平时${course.PSCJXS}% 期末${course.QMCJXS}%`);
                    } else {
                        // 未获取到系数，准备推算
                        course.PSCJXS = '?';  // '?' 表示待计算
                        course.QMCJXS = '?';
                        course._pscjxsNum = null;
                        course._qmcjxsNum = null;
                        course._coefficientsSource = 'unknown';
                        course._coefficientsInferred = false;
                        course._needPscj = true;
                        course._needQmcj = true;
                    }
                    
                    // 保存原始总成绩用于后续推算系数
                    course._originalZCJ = course.ZCJ;
                    
                    if (course._needPscj) needPscjCount++;
                    if (course._needQmcj) needQmcjCount++;
                    
                    courseMap.set(key, course);
                });

                console.log(`[深大成绩查询] 需要查询平时成绩: ${needPscjCount} 门, 期末成绩: ${needQmcjCount} 门`);

                let pscjFoundCount = 0;
                let qmcjFoundCount = 0;
                
                setQueryProgress(20, '正在查询详细成绩...', '正在并行扫描平时成绩和期末成绩。');

                // 3. 十线程并行分段查询策略
                // 10个线程分别处理10个分数段，每个线程处理约10个分数
                const scoreRanges = [
                    { start: 100, end: 91, label: '分段91-100' },
                    { start: 90, end: 81, label: '分段81-90' },
                    { start: 80, end: 71, label: '分段71-80' },
                    { start: 70, end: 61, label: '分段61-70' },
                    { start: 60, end: 51, label: '分段51-60' },
                    { start: 50, end: 41, label: '分段41-50' },
                    { start: 40, end: 31, label: '分段31-40' },
                    { start: 30, end: 21, label: '分段21-30' },
                    { start: 20, end: 11, label: '分段11-20' },
                    { start: 10, end: 0, label: '分段0-10' }
                ];
                
                // 共享状态（用于跟踪进度和提前终止）
                const sharedState = {
                    pscjFoundCount: 0,
                    qmcjFoundCount: 0,
                    queriedScores: new Set(),
                    allDone: false
                };
                
                // 更新进度显示
                const updateProgress = () => {
                    const totalScores = 101;
                    const scanProgress = Math.min((sharedState.queriedScores.size / totalScores) * 100, 100);
                    const progress = Math.min(20 + scanProgress * 0.78, 98);
                    setQueryProgress(
                        progress,
                        `并行查询中... [平时:${sharedState.pscjFoundCount}/${needPscjCount} 期末:${sharedState.qmcjFoundCount}/${needQmcjCount}] (已查${sharedState.queriedScores.size}个分数)`,
                        `已扫描 ${sharedState.queriedScores.size}/${totalScores} 个分数点。`
                    );
                };
                
                // 检查是否所有成绩都已找到
                const checkAllDone = () => {
                    if (sharedState.pscjFoundCount >= needPscjCount && sharedState.qmcjFoundCount >= needQmcjCount) {
                        sharedState.allDone = true;
                        return true;
                    }
                    return false;
                };
                
                // 尝试推算课程系数的函数（支持0:100情况）
                const tryInferCourseCoefficients = (course, scoreType, score) => {
                    // 如果已经是官方系数或已经推算过，则跳过
                    if (course._coefficientsSource === 'official' || course._coefficientsInferred) {
                        return;
                    }
                    
                    const zcj = course._originalZCJ;
                    if (zcj == null) {
                        return;
                    }
                    
                    // 快速检查：如果当前成绩等于总成绩，则为100:0或0:100
                    if (score === zcj) {
                        if (scoreType === 'PSCJ') {
                            // 平时成绩=总成绩，说明是100%平时成绩
                            course._pscjxsNum = 100;
                            course._qmcjxsNum = 0;
                            course.PSCJXS = '100*';
                            course.QMCJXS = '0*';
                            course.QMCJ = '-';  // 不需要期末成绩
                            course._needQmcj = false;
                            course._coefficientsInferred = true;
                            // 减少需要查询的期末成绩计数
                            if (sharedState.qmcjFoundCount < needQmcjCount) {
                                sharedState.qmcjFoundCount++;
                            }
                            console.log(`[系数推算] ${course.KCM}: 100%平时成绩 (平时=${score}=总成绩=${zcj})`);
                            renderResults();
                            return;
                        } else if (scoreType === 'QMCJ') {
                            // 期末成绩=总成绩，说明是100%期末成绩
                            course._pscjxsNum = 0;
                            course._qmcjxsNum = 100;
                            course.PSCJXS = '0*';
                            course.QMCJXS = '100*';
                            course.PSCJ = '-';  // 不需要平时成绩
                            course._needPscj = false;
                            course._coefficientsInferred = true;
                            // 减少需要查询的平时成绩计数
                            if (sharedState.pscjFoundCount < needPscjCount) {
                                sharedState.pscjFoundCount++;
                            }
                            console.log(`[系数推算] ${course.KCM}: 100%期末成绩 (期末=${score}=总成绩=${zcj})`);
                            renderResults();
                            return;
                        }
                    }
                    
                    // 检查是否两个成绩都已查到
                    const pscjStr = course.PSCJ;
                    const qmcjStr = course.QMCJ;
                    
                    if (pscjStr === 'N/A' || pscjStr === '-' || qmcjStr === 'N/A' || qmcjStr === '-') {
                        return; // 成绩未全部查到或不需要
                    }
                    
                    const pscj = parseFloat(pscjStr);
                    const qmcj = parseFloat(qmcjStr);
                    
                    if (isNaN(pscj) || isNaN(qmcj)) {
                        console.log(`[系数推算] ${course.KCM}: 数据不完整，无法推算`);
                        return;
                    }
                    
                    // 异步推算系数
                    setTimeout(() => {
                        const inferred = inferCoefficients(pscj, qmcj, zcj);
                        if (inferred) {
                            course._pscjxsNum = inferred.pscjxs;
                            course._qmcjxsNum = inferred.qmcjxs;
                            course.PSCJXS = String(inferred.pscjxs) + '*';
                            course.QMCJXS = String(inferred.qmcjxs) + '*';
                            course._coefficientsInferred = true;
                            console.log(`[系数推算] ${course.KCM}: 平时${inferred.pscjxs}% 期末${inferred.qmcjxs}%`);
                            
                            // 触发重新渲染
                            renderResults();
                        } else {
                            console.log(`[系数推算] ${course.KCM}: 无法推算系数 (平时=${pscj}, 期末=${qmcj}, 总成绩=${zcj})`);
                            course.PSCJXS = '?';
                            course.QMCJXS = '?';
                        }
                    }, 0);
                };
                
                // 单个分数段的查询任务
                const queryRangeTask = async (range) => {
                    console.log(`[深大成绩查询] 线程启动: ${range.label}`);
                    
                    for (let score = range.start; score >= range.end; score--) {
                        // 检查是否已全部完成
                        if (sharedState.allDone) {
                            console.log(`[深大成绩查询] ${range.label} 提前结束（所有成绩已找到）`);
                            break;
                        }
                        
                        // 标记该分数已查询
                        sharedState.queriedScores.add(score);
                        
                        // 查询平时成绩
                        if (sharedState.pscjFoundCount < needPscjCount) {
                            try {
                                const pscjRows = await performQuery(score, 'PSCJ');
                                pscjRows.forEach(row => {
                                    const key = row.KCM + row.XNXQDM_DISPLAY;
                                    const course = courseMap.get(key);
                                    if (course && course.PSCJ === 'N/A' && course._needPscj) {
                                        course.PSCJ = score.toString();
                                        sharedState.pscjFoundCount++;
                                        // 尝试推算系数（传入成绩类型和分数用于0:100判断）
                                        tryInferCourseCoefficients(course, 'PSCJ', score);
                                    }
                                });
                            } catch (e) {
                                console.error(`[深大成绩查询] ${range.label} 查询PSCJ=${score}失败:`, e);
                            }
                        }
                        
                        // 查询期末成绩
                        if (sharedState.qmcjFoundCount < needQmcjCount) {
                            try {
                                const qmcjRows = await performQuery(score, 'QMCJ');
                                qmcjRows.forEach(row => {
                                    const key = row.KCM + row.XNXQDM_DISPLAY;
                                    const course = courseMap.get(key);
                                    if (course && course.QMCJ === 'N/A' && course._needQmcj) {
                                        course.QMCJ = score.toString();
                                        sharedState.qmcjFoundCount++;
                                        // 尝试推算系数（传入成绩类型和分数用于0:100判断）
                                        tryInferCourseCoefficients(course, 'QMCJ', score);
                                    }
                                });
                            } catch (e) {
                                console.error(`[深大成绩查询] ${range.label} 查询QMCJ=${score}失败:`, e);
                            }
                        }
                        
                        // 更新数据和渲染
                        scriptState.courseData = Array.from(courseMap.values());
                        renderResults();
                        updateProgress();
                        
                        // 检查是否完成
                        checkAllDone();
                        
                        // 短暂延迟，避免请求过于密集
                        await new Promise(resolve => setTimeout(resolve, 30));
                    }
                    
                    console.log(`[深大成绩查询] ${range.label} 线程完成`);
                };
                
                // 启动10个并行线程
                console.log('[深大成绩查询] 启动10线程并行查询...');
                await Promise.all(scoreRanges.map(range => queryRangeTask(range)));
                
                // 更新最终计数
                pscjFoundCount = sharedState.pscjFoundCount;
                qmcjFoundCount = sharedState.qmcjFoundCount;

                setQueryProgress(100, `查询完成！共 ${courseMap.size} 门课程`, '结果已刷新到页面内表格和悬浮窗。', false);
                // 查询完成后隐藏进度条区域
                container.querySelector('.progress-container').classList.add('completed');
                exportBtn.disabled = false;

            } catch (err) {
                console.error("查询过程中发生错误:", err);
                setQueryProgress(100, `查询异常: ${err.message}`, '请检查登录状态或网络请求结果。', false);
            } finally {
                scriptState.isRunning = false;
                startBtn.disabled = false;
                renderInlineScorePanel();
            }
        });

        exportBtn.addEventListener('click', () => {
            if (scriptState.courseData.length === 0) {
                alert('没有成绩数据可导出。');
                return;
            }

            // 准备表头（与前端展示的数据一致，增加系数来源列）
            const header = [
                '学期', '课程号', '课程名称', '课程类别', '开课学院', '课程学分',
                '平时成绩', '平时系数(%)', '期末成绩', '期末系数(%)',
                '总成绩', '等级', '等级制成绩', '系数来源'
            ];

            // 准备数据行
            const dataRows = scriptState.courseData.map(course => {
                const { finalScore, grade } = calculateFinalScoreAndGrade(course);
                // 判断系数来源
                let coefficientSource = '未知';
                if (course._coefficientsSource === 'official') {
                    coefficientSource = '接口返回';
                } else if (course._coefficientsInferred) {
                    coefficientSource = '推算';
                } else if (course.PSCJXS && !course.PSCJXS.endsWith('*') && course.PSCJXS !== '?') {
                    coefficientSource = '接口返回(旧)';
                }
                
                return [
                    course.XNXQDM_DISPLAY || 'N/A',
                    course.KCH || 'N/A',
                    course.KCM || 'N/A',
                    course.KCLBDM_DISPLAY || 'N/A',
                    course.KKDWDM_DISPLAY || 'N/A',
                    course.XF || 'N/A',
                    course.PSCJ,
                    course.PSCJXS ? course.PSCJXS.replace('*', '') : 'N/A',
                    course.QMCJ,
                    course.QMCJXS ? course.QMCJXS.replace('*', '') : 'N/A',
                    finalScore,
                    grade,
                    course.XFJD || 'N/A',
                    coefficientSource
                ];
            });

            // 创建工作表数据（包含表头）
            const wsData = [header, ...dataRows];

            // 创建工作表
            const ws = XLSX.utils.aoa_to_sheet(wsData);

            // 设置列宽
            ws['!cols'] = [
                { wch: 22.5 },  // 学期
                { wch: 11 },    // 课程号
                { wch: 25 },    // 课程名称
                { wch: 12 },    // 课程类别
                { wch: 20 },    // 开课学院
                { wch: 10 },    // 课程学分
                { wch: 10 },    // 平时成绩
                { wch: 12 },    // 平时系数
                { wch: 10 },    // 期末成绩
                { wch: 12 },    // 期末系数
                { wch: 10 },    // 总成绩
                { wch: 8 },     // 等级
                { wch: 12 },    // 等级制成绩
                { wch: 10 }     // 系数来源
            ];

            // 创建工作簿
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '成绩单');

            // 生成文件名
            let filename = '深大详细成绩单.xlsx';
            if (scriptState.studentId && scriptState.studentName) {
                filename = `深大详细成绩单-${scriptState.studentId}-${scriptState.studentName}.xlsx`;
            }

            // 导出文件
            XLSX.writeFile(wb, filename);
        });
    }

    function calculateFinalScoreAndGrade(course) {
        // 使用内部存储的数值系数，处理系数未知的情况
        const pscjxs = course._pscjxsNum;
        const qmcjxs = course._qmcjxsNum;
        
        // 判断系数是否已知
        const pscjxsKnown = pscjxs !== null && pscjxs !== undefined;
        const qmcjxsKnown = qmcjxs !== null && qmcjxs !== undefined;
        
        // 解析成绩，'-' 表示不需要该成绩
        const pscjStr = course.PSCJ;
        const qmcjStr = course.QMCJ;
        const pscj = pscjStr === '-' ? null : parseFloat(pscjStr);
        const qmcj = qmcjStr === '-' ? null : parseFloat(qmcjStr);
        
        // 检查成绩是否已获取
        const hasPscj = pscjStr !== '-' && pscjStr !== 'N/A' && !isNaN(pscj);
        const hasQmcj = qmcjStr !== '-' && qmcjStr !== 'N/A' && !isNaN(qmcj);

        let rawFinalScore;

        // 情况1：系数都未知，无法计算，使用服务器返回的总成绩
        if (!pscjxsKnown && !qmcjxsKnown) {
            if (course.ZCJ != null) {
                return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
            }
            // 如果两个成绩都已获取，尝试简单平均（仅作为备选）
            if (hasPscj && hasQmcj) {
                rawFinalScore = (pscj + qmcj) / 2;
            } else {
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        // 情况2：只有平时成绩系数有效（期末系数为0或未知）
        else if (pscjxsKnown && pscjxs === 100) {
            if (hasPscj) {
                rawFinalScore = pscj;
            } else {
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        else if (pscjxsKnown && pscjxs > 0 && qmcjxsKnown && qmcjxs === 0) {
            if (hasPscj) {
                rawFinalScore = pscj;
            } else {
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        // 情况3：只有期末成绩系数有效（平时系数为0或未知）
        else if (qmcjxsKnown && qmcjxs === 100) {
            if (hasQmcj) {
                rawFinalScore = qmcj;
            } else {
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        else if (qmcjxsKnown && qmcjxs > 0 && pscjxsKnown && pscjxs === 0) {
            if (hasQmcj) {
                rawFinalScore = qmcj;
            } else {
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        // 情况4：正常情况，两个系数都有效且都 > 0
        else if (pscjxsKnown && qmcjxsKnown && pscjxs > 0 && qmcjxs > 0) {
            if (hasPscj && hasQmcj) {
                rawFinalScore = (pscj * pscjxs / 100) + (qmcj * qmcjxs / 100);
            } else {
                // 成绩不完整，使用服务器返回的总成绩
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        // 其他情况：使用服务器返回的总成绩
        else {
            if (course.ZCJ != null) {
                return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
            }
            return { finalScore: 'N/A', grade: 'N/A' };
        }

        const finalScore = Math.round(rawFinalScore);
        let grade = 'F';
        if (finalScore >= 93) grade = 'A+';
        else if (finalScore >= 85) grade = 'A';
        else if (finalScore >= 80) grade = 'B+';
        else if (finalScore >= 75) grade = 'B';
        else if (finalScore >= 70) grade = 'C+';
        else if (finalScore >= 65) grade = 'C';
        else if (finalScore >= 60) grade = 'D';

        return { finalScore, grade };
    }

    function calculateGPA(courses) {
        let totalPoints = 0;
        let totalCredits = 0;
        courses.forEach(course => {
            const credit = parseFloat(course.XF);
            const point = parseFloat(course.XFJD);
            if (!isNaN(credit) && !isNaN(point)) {
                totalPoints += credit * point;
                totalCredits += credit;
            }
        });
        return totalCredits > 0 ? (totalPoints / totalCredits).toFixed(2) : '0.00';
    }

    // 渲染 GPA 趋势折线图
    function renderGPAChart(semesterData, yearData) {
        if (semesterData.length < 2 && yearData.length < 2) {
            return ''; // 数据点太少，不显示图表
        }

        const chartWidth = 440;
        const chartHeight = 120;
        const padding = { top: 20, right: 30, bottom: 30, left: 35 };
        const innerWidth = chartWidth - padding.left - padding.right;
        const innerHeight = chartHeight - padding.top - padding.bottom;

        // 生成单个折线图的 SVG
        function generateLineChart(data, color, title) {
            if (data.length < 2) return '';
            
            const gpas = data.map(d => d.gpa);
            const minGPA = Math.max(0, Math.floor(Math.min(...gpas) * 10) / 10 - 0.2);
            const maxGPA = Math.min(5, Math.ceil(Math.max(...gpas) * 10) / 10 + 0.2);
            const gpaRange = maxGPA - minGPA || 1;

            // 计算点的位置
            const points = data.map((d, i) => {
                const x = padding.left + (i / (data.length - 1)) * innerWidth;
                const y = padding.top + innerHeight - ((d.gpa - minGPA) / gpaRange) * innerHeight;
                return { x, y, gpa: d.gpa, label: d.label || d.year };
            });

            // 生成折线路径
            const linePath = points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ');
            
            // 生成填充区域路径
            const areaPath = linePath + ` L ${points[points.length - 1].x} ${padding.top + innerHeight} L ${points[0].x} ${padding.top + innerHeight} Z`;

            // Y轴刻度
            const yTicks = [];
            const tickCount = 4;
            for (let i = 0; i <= tickCount; i++) {
                const val = minGPA + (gpaRange * i / tickCount);
                const y = padding.top + innerHeight - (i / tickCount) * innerHeight;
                yTicks.push({ val: val.toFixed(1), y });
            }

            return `
                <div class="score-chart-card">
                    <div style="font-size:0.8rem;color:#666;margin-bottom:4px;font-weight:500;">${title}</div>
                    <svg width="${chartWidth}" height="${chartHeight}" viewBox="0 0 ${chartWidth} ${chartHeight}" style="background:#fff;border-radius:6px;border:1px solid #e0e0e0;">
                        <!-- 网格线 -->
                        ${yTicks.map(t => `<line x1="${padding.left}" y1="${t.y}" x2="${chartWidth - padding.right}" y2="${t.y}" stroke="#f0f0f0" stroke-width="1"/>`).join('')}
                        
                        <!-- Y轴刻度值 -->
                        ${yTicks.map(t => `<text x="${padding.left - 5}" y="${t.y + 3}" text-anchor="end" font-size="10" fill="#999">${t.val}</text>`).join('')}
                        
                        <!-- X轴标签 -->
                        ${points.map((p, i) => `<text x="${p.x}" y="${chartHeight - 8}" text-anchor="middle" font-size="9" fill="#666">${p.label}</text>`).join('')}
                        
                        <!-- 填充区域 -->
                        <path d="${areaPath}" fill="${color}" fill-opacity="0.1"/>
                        
                        <!-- 折线 -->
                        <path d="${linePath}" fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                        
                        <!-- 数据点 -->
                        ${points.map(p => `
                            <circle cx="${p.x}" cy="${p.y}" r="4" fill="#fff" stroke="${color}" stroke-width="2"/>
                            <text x="${p.x}" y="${p.y - 8}" text-anchor="middle" font-size="10" font-weight="bold" fill="${color}">${p.gpa.toFixed(2)}</text>
                        `).join('')}
                    </svg>
                </div>
            `;
        }

        let html = '<div class="score-chart-grid">';
        
        // 学期 GPA 趋势
        if (semesterData.length >= 2) {
            html += generateLineChart(semesterData, '#1976d2', '📈 学期 GPA 趋势');
        }
        
        // 学年 GPA 趋势
        if (yearData.length >= 2) {
            const yearChartData = yearData.slice().reverse().map(d => ({ label: d.year, gpa: parseFloat(d.gpa) }));
            html += generateLineChart(yearChartData, '#43a047', '📊 学年 GPA 趋势');
        }
        
        html += '</div>';
        return html;
    }

    function appendScoreResultsContent(container, courses) {
        container.innerHTML = '';

        if (!courses || courses.length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:20px;color:#999;">暂无数据</div>';
            return;
        }

        appendScoreSummary(container, courses);
        appendScoreSemesterTables(container, buildSemesterGroups(courses));
    }

    function appendScoreSummary(container, courses) {
        const totalGPA = calculateGPA(courses);

        const yearGroups = {};
        courses.forEach(course => {
            const year = course.XNXQDM ? course.XNXQDM.substring(0, 9) : '未知学年';
            if (!yearGroups[year]) yearGroups[year] = [];
            yearGroups[year].push(course);
        });
        
        const yearGPAs = Object.keys(yearGroups).sort().reverse().map(year => {
            return { year, gpa: calculateGPA(yearGroups[year]) };
        });

        // 3. 计算学期 GPA
        const semesterGPAData = [];
        const semesterKeys = [...new Set(courses.map(c => c.XNXQDM))].sort();
        semesterKeys.forEach(xnxqdm => {
            const semesterCourses = courses.filter(c => c.XNXQDM === xnxqdm);
            const displayName = semesterCourses[0]?.XNXQDM_DISPLAY || xnxqdm;
            semesterGPAData.push({
                key: xnxqdm,
                label: displayName.replace('学年', '').replace('学期', ''),
                gpa: parseFloat(calculateGPA(semesterCourses))
            });
        });

        const summaryDiv = document.createElement('div');
        summaryDiv.className = 'score-summary-card';
        let summaryHTML = `<div class="score-summary-title">总 GPA: ${totalGPA}</div>`;
        summaryHTML += '<div class="score-summary-grid">';
        summaryHTML += buildScoreSummaryTable(
            '学期 GPA',
            semesterGPAData.slice().reverse().map(item => ({
                label: item.label,
                value: item.gpa.toFixed(2)
            }))
        );
        summaryHTML += buildScoreSummaryTable(
            '学年 GPA',
            yearGPAs.map(item => ({
                label: `${item.year}学年`,
                value: item.gpa
            }))
        );
        summaryHTML += '</div>';
        summaryHTML += renderGPAChart(semesterGPAData, yearGPAs);
        summaryDiv.innerHTML = summaryHTML;
        container.appendChild(summaryDiv);
    }

    function buildScoreSummaryTable(title, rows) {
        const body = rows.length
            ? rows.map(row => `
                <tr>
                    <th>${escapeHtml(row.label)}</th>
                    <td>${escapeHtml(row.value)}</td>
                </tr>
            `).join('')
            : '<tr><th>暂无数据</th><td>-</td></tr>';

        return `
            <div class="score-summary-panel">
                <div class="score-summary-panel-title">${escapeHtml(title)}</div>
                <table class="score-summary-table">
                    <tbody>${body}</tbody>
                </table>
            </div>
        `;
    }

    function escapeHtml(value) {
        return String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function buildSemesterGroups(courses) {
        const sortedCourses = [...courses].sort((a, b) => {
            if (a.XNXQDM !== b.XNXQDM) {
                return (b.XNXQDM || '').localeCompare(a.XNXQDM || '');
            }
            return String(a.KCM || '').localeCompare(String(b.KCM || ''), 'zh-Hans');
        });

        const semesterGroups = new Map();
        sortedCourses.forEach(course => {
            const key = course.XNXQDM_DISPLAY || course.XNXQDM || '未知学期';
            if (!semesterGroups.has(key)) {
                semesterGroups.set(key, []);
            }
            semesterGroups.get(key).push(course);
        });

        return semesterGroups;
    }

    function getScoreTableColumns() {
        return [
            {
                title: '课程名称',
                field: 'courseName',
                type: 'text',
                colClass: 'score-col-course',
                value: course => course.KCM
            },
            {
                title: '课程性质',
                field: 'nature',
                type: 'text',
                colClass: 'score-col-nature',
                value: course => formatCourseNature(course)
            },
            {
                title: '学分',
                field: 'credit',
                type: 'number',
                colClass: 'score-col-credit',
                value: course => course.XF
            },
            {
                title: '总成绩',
                field: 'totalScore',
                type: 'number',
                colClass: 'score-col-total',
                value: course => calculateFinalScoreAndGrade(course).finalScore
            },
            {
                title: '等级成绩',
                field: 'grade',
                type: 'text',
                colClass: 'score-col-grade',
                value: course => {
                    const { grade } = calculateFinalScoreAndGrade(course);
                    return course.DJCJMC || grade;
                }
            },
            {
                title: '平时成绩',
                field: 'regularScore',
                type: 'number',
                colClass: 'score-col-regular',
                value: course => course.PSCJ
            },
            {
                title: '期末成绩',
                field: 'finalScorePart',
                type: 'number',
                colClass: 'score-col-final',
                value: course => course.QMCJ
            },
            {
                title: '平时成绩系数',
                field: 'regularCoeff',
                type: 'number',
                colClass: 'score-col-regular-coeff',
                value: course => course.PSCJXS
            },
            {
                title: '期末成绩系数',
                field: 'finalCoeff',
                type: 'number',
                colClass: 'score-col-final-coeff',
                value: course => course.QMCJXS
            }
        ];
    }

    function getCurrentTableSort() {
        const defaultSort = { field: 'courseName', direction: 'asc' };
        const fields = getScoreTableColumns().map(column => column.field);
        const sort = scriptState.tableSort || defaultSort;
        const direction = sort.direction === 'desc' ? 'desc' : 'asc';

        if (!fields.includes(sort.field)) {
            return defaultSort;
        }

        return {
            field: sort.field,
            direction
        };
    }

    function setTableSort(field) {
        const currentSort = getCurrentTableSort();
        const nextDirection = currentSort.field === field && currentSort.direction === 'asc' ? 'desc' : 'asc';

        scriptState.tableSort = {
            field,
            direction: nextDirection
        };

        renderInlineScorePanel();
    }

    function appendScoreTableHeaderCell(row, column) {
        const currentSort = getCurrentTableSort();
        const isActive = currentSort.field === column.field;
        const cell = appendTableCell(row, '', 'th', isActive ? 'sortable active-sort' : 'sortable');
        const label = document.createElement('span');
        label.textContent = column.title;

        const indicator = document.createElement('span');
        indicator.className = 'sort-indicator';
        indicator.textContent = isActive ? (currentSort.direction === 'asc' ? '↑' : '↓') : '';

        cell.setAttribute('aria-sort', isActive ? (currentSort.direction === 'asc' ? 'ascending' : 'descending') : 'none');
        cell.title = `点击按${column.title}排序`;
        cell.addEventListener('click', () => setTableSort(column.field));
        cell.appendChild(label);
        cell.appendChild(indicator);
        return cell;
    }

    function sortCoursesForTable(courses) {
        const currentSort = getCurrentTableSort();
        const columns = getScoreTableColumns();
        const column = columns.find(item => item.field === currentSort.field) || columns[0];

        return [...courses].sort((a, b) => compareScoreTableCourses(a, b, column, currentSort.direction));
    }

    function compareScoreTableCourses(a, b, column, direction) {
        const left = getTableSortValue(a, column);
        const right = getTableSortValue(b, column);

        if (left.missing && right.missing) {
            return compareCourseNameAsc(a, b);
        }
        if (left.missing) return 1;
        if (right.missing) return -1;

        const primary = column.type === 'number'
            ? left.value - right.value
            : compareSortText(left.value, right.value);

        if (primary !== 0) {
            return direction === 'desc' ? -primary : primary;
        }

        return compareCourseNameAsc(a, b);
    }

    function getTableSortValue(course, column) {
        const rawValue = column.value(course);

        if (column.type === 'number') {
            const value = parseSortableNumber(rawValue);
            return {
                missing: value === null,
                value
            };
        }

        const value = normalizeSortText(rawValue);
        return {
            missing: value === '',
            value
        };
    }

    function compareCourseNameAsc(a, b) {
        const courseName = compareSortText(normalizeSortText(a.KCM), normalizeSortText(b.KCM));
        if (courseName !== 0) return courseName;
        return compareSortText(normalizeSortText(a.KCH || a.KCDM), normalizeSortText(b.KCH || b.KCDM));
    }

    function compareSortText(left, right) {
        return String(left || '').localeCompare(String(right || ''), 'zh-Hans', {
            numeric: true,
            sensitivity: 'base'
        });
    }

    function normalizeSortText(value) {
        if (value === null || value === undefined) return '';
        const text = String(value).trim();
        if (!text || text === '-' || text === 'N/A' || text === '?' || text === '查询中') return '';
        return text;
    }

    function parseSortableNumber(value) {
        if (value === null || value === undefined) return null;
        const text = String(value).replace(/\*/g, '').replace(/%/g, '').trim();
        if (!text || text === '-' || text === 'N/A' || text === '?' || text === '查询中') return null;

        const matched = text.match(/-?\d+(?:\.\d+)?/);
        if (!matched) return null;

        const numeric = Number(matched[0]);
        return Number.isFinite(numeric) ? numeric : null;
    }

    function appendScoreSemesterTables(container, semesterGroups) {
        semesterGroups.forEach((semesterCourses, semesterName) => {
            const semesterGPA = calculateGPA(semesterCourses);

            const section = document.createElement('section');
            section.className = 'score-semester-section';

            const semesterHeader = document.createElement('div');
            semesterHeader.className = 'score-semester-header';
            semesterHeader.innerHTML = `<h4>${semesterName}</h4><span>GPA: ${semesterGPA}</span>`;
            section.appendChild(semesterHeader);

            const tableWrap = document.createElement('div');
            tableWrap.className = 'score-table-wrap';
            tableWrap.appendChild(createScoreTable(semesterCourses));
            section.appendChild(tableWrap);
            container.appendChild(section);
        });
    }

    function createScoreTable(courses) {
        const table = document.createElement('table');
        table.className = 'score-table';
        const columns = getScoreTableColumns();
        const sortedCourses = sortCoursesForTable(courses);

        const colGroup = document.createElement('colgroup');
        columns.forEach(column => {
            const col = document.createElement('col');
            col.className = column.colClass;
            colGroup.appendChild(col);
        });
        table.appendChild(colGroup);

        const thead = document.createElement('thead');
        const headRow = document.createElement('tr');
        columns.forEach(column => appendScoreTableHeaderCell(headRow, column));
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        sortedCourses.forEach(course => {
            const { finalScore, grade } = calculateFinalScoreAndGrade(course);
            const tr = document.createElement('tr');

            appendTableCell(tr, normalizeDisplayValue(course.KCM), 'td', 'course-name-cell');
            appendTableCell(tr, formatCourseNature(course), 'td');
            appendTableCell(tr, normalizeDisplayValue(course.XF), 'td');
            appendTableCell(tr, normalizeDisplayValue(finalScore), 'td', 'score-number');
            appendTableCell(tr, normalizeDisplayValue(course.DJCJMC || grade), 'td');
            appendTableCell(tr, formatScoreDisplay(course.PSCJ), 'td', scoreCellClass(course.PSCJ));
            appendTableCell(tr, formatScoreDisplay(course.QMCJ), 'td', scoreCellClass(course.QMCJ));
            appendTableCell(tr, formatCoefficient(course.PSCJXS), 'td', 'score-coeff');
            appendTableCell(tr, formatCoefficient(course.QMCJXS), 'td', 'score-coeff');

            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        return table;
    }

    function appendTableCell(row, value, tagName, className) {
        const cell = document.createElement(tagName);
        cell.textContent = value;
        if (className) cell.className = className;
        row.appendChild(cell);
        return cell;
    }

    function normalizeDisplayValue(value) {
        if (value === null || value === undefined || value === '') return '-';
        return String(value);
    }

    function formatScoreDisplay(value) {
        if (value === 'N/A') return '查询中';
        return normalizeDisplayValue(value);
    }

    function scoreCellClass(value) {
        return value === 'N/A' || value === '?' || value === null || value === undefined ? 'score-muted' : 'score-number';
    }

    function formatCourseNature(course) {
        const rawValue = normalizeDisplayValue(course?.KCXZDM_DISPLAY || course?.KCXZDM);
        const code = String(course?.KCXZDM || course?.KCXZDM_DISPLAY || '').trim();
        const natureMap = {
            '01': '必修课',
            '02': '选修课',
            '必修': '必修课',
            '选修': '选修课'
        };

        if (natureMap[code]) return natureMap[code];
        if (natureMap[rawValue]) return natureMap[rawValue];
        return rawValue;
    }

    function renderResults() {
        const resultsEl = scriptState.container.querySelector('#score-results');
        renderFloatingResults(resultsEl, scriptState.courseData);
        renderInlineScorePanel();
    }

    function renderFloatingResults(resultsEl, courses) {
        resultsEl.innerHTML = '';

        if (!courses || courses.length === 0) {
            resultsEl.innerHTML = '<div style="text-align:center;padding:20px;color:#999;">暂无数据</div>';
            return;
        }

        const totalGPA = calculateGPA(courses);

        const yearGroups = {};
        courses.forEach(course => {
            const year = course.XNXQDM ? course.XNXQDM.substring(0, 9) : '未知学年';
            if (!yearGroups[year]) yearGroups[year] = [];
            yearGroups[year].push(course);
        });

        const yearGPAs = Object.keys(yearGroups).sort().reverse().map(year => {
            return { year, gpa: calculateGPA(yearGroups[year]) };
        });

        const semesterGPAData = [];
        const semesterKeys = [...new Set(courses.map(c => c.XNXQDM))].sort();
        semesterKeys.forEach(xnxqdm => {
            const semesterCourses = courses.filter(c => c.XNXQDM === xnxqdm);
            const displayName = semesterCourses[0]?.XNXQDM_DISPLAY || xnxqdm;
            semesterGPAData.push({
                key: xnxqdm,
                label: displayName.replace('学年', '').replace('学期', ''),
                gpa: parseFloat(calculateGPA(semesterCourses))
            });
        });

        const summaryDiv = document.createElement('div');
        summaryDiv.style.cssText = 'background:#e3f2fd;padding:12px;border-radius:8px;margin-bottom:16px;border:1px solid #bbdefb;';
        let summaryHTML = `<div style="font-size:1.1rem;font-weight:bold;color:#1565c0;margin-bottom:8px;">总 GPA: ${totalGPA}</div>`;
        summaryHTML += '<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px;">';
        yearGPAs.forEach(item => {
            summaryHTML += `<span style="background:#fff;padding:4px 8px;border-radius:4px;font-size:0.85rem;color:#555;border:1px solid #e0e0e0;">${item.year}学年: <b>${item.gpa}</b></span>`;
        });
        summaryHTML += '</div>';
        summaryHTML += renderGPAChart(semesterGPAData, yearGPAs);
        summaryDiv.innerHTML = summaryHTML;
        resultsEl.appendChild(summaryDiv);

        const semesterGroups = buildSemesterGroups(courses);
        semesterGroups.forEach((semesterCourses, semesterName) => {
            const semesterGPA = calculateGPA(semesterCourses);

            const semesterHeader = document.createElement('div');
            semesterHeader.style.cssText = 'margin:12px 0 8px 0;padding:8px 0 4px 0;border-bottom:2px solid #eee;display:flex;justify-content:space-between;align-items:center;position:sticky;top:-4px;background:#f9f9f9;z-index:10;';
            semesterHeader.innerHTML = `<h4 style="margin:0;color:#333;">${semesterName}</h4><span style="font-weight:bold;color:#4caf50;">GPA: ${semesterGPA}</span>`;
            resultsEl.appendChild(semesterHeader);

            semesterCourses.forEach(course => {
                const { finalScore, grade } = calculateFinalScoreAndGrade(course);
                const item = document.createElement('div');
                item.className = 'course-item';
                item.innerHTML = `
                    <div class="course-header">
                        <strong>${course.KCM}</strong>
                        <span>${course.KCLBDM_DISPLAY || ''}</span>
                    </div>

                    <div class="course-detail">
                        <span class="tag">课程学分: ${course.XF || 'N/A'}</span>
                        <span class="tag">等级制成绩: ${course.XFJD || 'N/A'}</span>
                    </div>
                    <div class="course-detail">
                        开课学院: ${course.KKDWDM_DISPLAY || 'N/A'}
                    </div>

                    <div class="course-detail full-width score-row">
                        <span>平时: <b style="color: #4CAF50;">${course.PSCJ}</b> (${formatCoefficient(course.PSCJXS)})</span>
                        <span>期末: <b style="color: #FF5722;">${course.QMCJ}</b> (${formatCoefficient(course.QMCJXS)})</span>
                    </div>

                    <div class="course-detail full-width score-row" style="margin-top: 4px; padding-top: 4px; border-top: 1px solid #eee;">
                        <span>总评: <span class="final-score">${finalScore}</span> <span class="final-score">(${grade})</span></span>
                    </div>
                `;
                resultsEl.appendChild(item);
            });
        });
    }

    // 格式化系数显示
    function formatCoefficient(xs) {
        if (xs === null || xs === undefined || xs === '' || xs === '?') return '?';
        if (xs === '-') return '-';
        if (typeof xs === 'string' && xs.endsWith('*')) {
            // 推断值，显示带提示
            return xs.replace('*', '') + '% (推断)';
        }
        return xs + '%';
    }

    function installInlineScoreTab() {
        if (scriptState.inlineScoreTab.installed) return;

        waitForElement('.jqx-tabs-title-container', 12000).then(tabList => {
            if (!tabList || scriptState.inlineScoreTab.installed) return;
            if (document.querySelector('[data-szu-score-inline-tab="1"]')) return;

            const tab = createInlineScoreTab();
            const panel = createInlineScorePanel();
            const contentContainer = findInlineScoreContentContainer();

            if (!contentContainer) {
                console.warn('[深大成绩查询] 未找到官方成绩页面内容容器，无法注入详细成绩表格页。');
                return;
            }

            tabList.appendChild(tab);
            contentContainer.appendChild(panel);

            tab.addEventListener('click', () => activateInlineScoreTab(tab, panel));
            bindOriginalTabsForInlinePanel(tab, panel);

            scriptState.inlineScoreTab.installed = true;
            scriptState.inlineScoreTab.tab = tab;
            scriptState.inlineScoreTab.panel = panel;
            renderInlineScorePanel();
        });
    }

    function waitForElement(selector, timeoutMs) {
        const existing = document.querySelector(selector);
        if (existing) return Promise.resolve(existing);

        return new Promise(resolve => {
            const started = Date.now();
            const timer = setInterval(() => {
                const element = document.querySelector(selector);
                if (element) {
                    clearInterval(timer);
                    resolve(element);
                    return;
                }

                if (Date.now() - started >= timeoutMs) {
                    clearInterval(timer);
                    resolve(null);
                }
            }, 250);
        });
    }

    function createInlineScoreTab() {
        const tab = document.createElement('li');
        tab.setAttribute('role', 'tab');
        tab.setAttribute('data-szu-score-inline-tab', '1');
        tab.className = 'jqx-reset jqx-disableselect jqx-tabs-title jqx-item jqx-rc-t jqx-fill-state-pressed';
        tab.style.float = 'left';

        const titleWrapper = document.createElement('div');
        titleWrapper.className = 'jqx-tabs-titleWrapper';
        titleWrapper.style.cssText = 'outline:none;position:relative;z-index:15;height:100%;';

        const titleContentWrapper = document.createElement('div');
        titleContentWrapper.className = 'jqx-tabs-titleContentWrapper jqx-disableselect';
        titleContentWrapper.style.cssText = 'float:left;margin-top:-0.5px;';
        titleContentWrapper.textContent = '详细成绩表格版';

        titleWrapper.appendChild(titleContentWrapper);
        tab.appendChild(titleWrapper);
        return tab;
    }

    function createInlineScorePanel() {
        const panel = document.createElement('div');
        panel.className = 'cjcx-tab-content-2 bh-mt-8 jqx-tabs-content-element jqx-rc-b szu-inline-score-panel';
        panel.setAttribute('role', 'tabpanel');
        panel.setAttribute('data-szu-score-inline-panel', '1');
        panel.style.display = 'none';
        return panel;
    }

    function findInlineScoreContentContainer() {
        const existingPanel = document.querySelector('.jqx-tabs-content-element');
        if (existingPanel && existingPanel.parentElement) return existingPanel.parentElement;

        return document.querySelector('.jqx-widget-content') || document.querySelector('[role="tabpanel"]')?.parentElement;
    }

    function bindOriginalTabsForInlinePanel(inlineTab, inlinePanel) {
        document.querySelectorAll('ul.jqx-tabs-title-container > li').forEach(tab => {
            if (tab === inlineTab) return;
            tab.addEventListener('click', () => {
                inlineTab.classList.remove('jqx-tabs-title-selected-top');
                inlinePanel.style.display = 'none';
            });
        });
    }

    function activateInlineScoreTab(tab, panel) {
        document.querySelectorAll('.jqx-tabs-title-container > li').forEach(item => {
            item.classList.remove('jqx-tabs-title-selected-top');
        });
        tab.classList.add('jqx-tabs-title-selected-top');

        document.querySelectorAll('.jqx-tabs-content-element').forEach(content => {
            content.style.display = 'none';
        });
        panel.style.display = 'block';
        renderInlineScorePanel();
    }

    function renderInlineScorePanel() {
        const panel = scriptState.inlineScoreTab.panel;
        if (!panel) return;

        panel.innerHTML = '';

        const toolbar = document.createElement('div');
        toolbar.className = 'szu-inline-score-toolbar';

        const refreshButton = document.createElement('button');
        refreshButton.type = 'button';
        refreshButton.textContent = scriptState.courseData.length ? '重新获取成绩' : '获取详细成绩';
        refreshButton.disabled = scriptState.isRunning;
        refreshButton.addEventListener('click', () => {
            if (scriptState.isRunning) return;
            scriptState.container?.querySelector('#start-query')?.click();
            renderInlineScorePanel();
        });

        const openPanelButton = document.createElement('button');
        openPanelButton.type = 'button';
        openPanelButton.textContent = '打开悬浮窗';
        openPanelButton.addEventListener('click', () => {
            scriptState.container?.classList.remove('hidden');
        });

        const hint = document.createElement('span');
        hint.className = 'szu-inline-score-hint';
        hint.textContent = scriptState.isRunning
            ? '正在查询，结果会自动刷新到这里。'
            : '使用当前助手查询结果渲染，与悬浮窗和 Excel 导出保持一致。';

        toolbar.appendChild(refreshButton);
        toolbar.appendChild(openPanelButton);
        toolbar.appendChild(hint);
        panel.appendChild(toolbar);

        const progressCard = createInlineProgressCard();
        if (progressCard) {
            panel.appendChild(progressCard);
        }

        const resultHost = document.createElement('div');
        panel.appendChild(resultHost);

        if (scriptState.courseData.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'szu-inline-score-empty';
            empty.textContent = scriptState.isRunning ? '正在获取成绩数据...' : '暂无成绩数据，请点击上方按钮开始查询。';
            resultHost.appendChild(empty);
            return;
        }

        appendScoreResultsContent(resultHost, scriptState.courseData);
    }

    function createInlineProgressCard() {
        const progress = scriptState.queryProgress;
        if (!progress?.updatedAt) return null;

        const percent = normalizeProgressPercent(progress.percent);
        const card = document.createElement('div');
        card.className = 'szu-inline-progress-card';

        const head = document.createElement('div');
        head.className = 'szu-inline-progress-head';

        const message = document.createElement('span');
        message.className = 'szu-inline-progress-message';
        message.textContent = progress.message || '准备就绪';

        const percentText = document.createElement('span');
        percentText.className = 'szu-inline-progress-percent';
        percentText.textContent = `${Math.round(percent)}%`;

        head.appendChild(message);
        head.appendChild(percentText);

        const track = document.createElement('div');
        track.className = 'szu-inline-progress-track';

        const fill = document.createElement('div');
        fill.className = 'szu-inline-progress-fill';
        fill.style.width = `${percent}%`;
        track.appendChild(fill);

        card.appendChild(head);
        card.appendChild(track);

        if (progress.detail) {
            const detail = document.createElement('div');
            detail.className = 'szu-inline-progress-detail';
            detail.textContent = progress.detail;
            card.appendChild(detail);
        }

        return card;
    }

    function setQueryProgress(percent, message, detail = '', active = true) {
        const normalizedPercent = normalizeProgressPercent(percent);
        scriptState.queryProgress = {
            active,
            percent: normalizedPercent,
            message,
            detail,
            updatedAt: new Date().toISOString()
        };

        const statusEl = scriptState.container?.querySelector('#status');
        const progressEl = scriptState.container?.querySelector('#progress');
        const progressContainer = scriptState.container?.querySelector('.progress-container');

        if (statusEl) statusEl.textContent = message;
        if (progressEl) progressEl.style.width = `${normalizedPercent}%`;
        if (progressContainer && active) {
            progressContainer.classList.remove('completed');
            progressContainer.classList.add('active');
        }

        updateInlineProgressCard();
    }

    function updateInlineProgressCard() {
        const panel = scriptState.inlineScoreTab.panel;
        if (!panel) return;

        let card = panel.querySelector('.szu-inline-progress-card');
        const progress = scriptState.queryProgress;

        if (!progress?.updatedAt) {
            if (card) card.remove();
            return;
        }

        if (!card) {
            renderInlineScorePanel();
            return;
        }

        const percent = normalizeProgressPercent(progress.percent);
        const message = card.querySelector('.szu-inline-progress-message');
        const percentText = card.querySelector('.szu-inline-progress-percent');
        const fill = card.querySelector('.szu-inline-progress-fill');

        if (message) message.textContent = progress.message || '准备就绪';
        if (percentText) percentText.textContent = `${Math.round(percent)}%`;
        if (fill) {
            requestAnimationFrame(() => {
                fill.style.width = `${percent}%`;
            });
        }

        let detail = card.querySelector('.szu-inline-progress-detail');
        if (progress.detail) {
            if (!detail) {
                detail = document.createElement('div');
                detail.className = 'szu-inline-progress-detail';
                card.appendChild(detail);
            }
            detail.textContent = progress.detail;
        } else if (detail) {
            detail.remove();
        }
    }

    function normalizeProgressPercent(percent) {
        const numeric = Number(percent);
        if (!Number.isFinite(numeric)) return 0;
        return Math.max(0, Math.min(100, numeric));
    }

    // 更新开发者模式数据显示
    function updateDevDataDisplay() {
        if (!scriptState.container) return;

        const initialDataEl = scriptState.container.querySelector('#dev-initial-data');

        if (initialDataEl && scriptState.rawData.initialCourses !== null) {
            initialDataEl.textContent = JSON.stringify(scriptState.rawData.initialCourses, null, 2);
        }

        updateDevQueryDisplay();
        updateDevProbeDisplay();
        updateDevNetworkDisplay();
    }

    // 更新轮询查询结果显示
    function updateDevQueryDisplay() {
        if (!scriptState.container) return;

        const queryListEl = scriptState.container.querySelector('#dev-query-list');
        const queryCountEl = scriptState.container.querySelector('#dev-query-count');

        if (!queryListEl || !queryCountEl) return;

        const results = scriptState.rawData.queryResults;
        queryCountEl.textContent = results.length;

        if (results.length === 0) {
            queryListEl.innerHTML = '<div style="padding:12px;color:#999;text-align:center;">暂无查询记录</div>';
            return;
        }

        // 只显示最近的100条记录，避免DOM过多
        const displayResults = results.slice(-100);

        queryListEl.innerHTML = displayResults.map((item, idx) => {
            const realIdx = results.length - displayResults.length + idx;
            const badgeClass = item.type === 'PSCJ' ? 'pscj' : 'qmcj';
            const typeLabel = item.type === 'PSCJ' ? '平时' : '期末';
            const rowCount = item.rows ? item.rows.length : 0;

            return `
                <div class="dev-query-item">
                    <div class="dev-query-header" onclick="this.nextElementSibling.classList.toggle('expanded')">
                        <span>#${realIdx + 1} 查询 ${typeLabel}=${item.score}</span>
                        <span>
                            <span class="dev-query-badge ${badgeClass}">${typeLabel}</span>
                            <span class="dev-query-badge count">${rowCount}条</span>
                        </span>
                    </div>
                    <div class="dev-query-body">${JSON.stringify(item, null, 2)}</div>
                </div>
            `;
        }).join('');
    }

    // 添加单条查询结果到记录
    function addQueryResult(score, type, rows, rawResponse) {
        const result = {
            timestamp: new Date().toISOString(),
            score: score,
            type: type,
            rowCount: rows.length,
            rows: rows,
            rawResponse: rawResponse
        };

        scriptState.rawData.queryResults.push(result);

        // 如果开发者模式开启，实时更新显示
        if (scriptState.devMode) {
            updateDevQueryDisplay();
        }
    }

    function updateDevProbeDisplay() {
        if (!scriptState.container) return;

        const probeDataEl = scriptState.container.querySelector('#dev-probe-data');
        const downloadBtn = scriptState.container.querySelector('#dev-download-probe-results');

        if (probeDataEl) {
            probeDataEl.textContent = scriptState.rawData.probeResults
                ? JSON.stringify(scriptState.rawData.probeResults, null, 2)
                : '暂无数据';
        }

        if (downloadBtn) {
            downloadBtn.disabled = !scriptState.rawData.probeResults;
        }
    }

    function formatDateTimeForFilename(date) {
        const pad = (value) => String(value).padStart(2, '0');
        return [
            date.getFullYear(),
            pad(date.getMonth() + 1),
            pad(date.getDate())
        ].join('') + '-' + [
            pad(date.getHours()),
            pad(date.getMinutes()),
            pad(date.getSeconds())
        ].join('');
    }

    function downloadJsonFile(data, filename) {
        const json = JSON.stringify(data, null, 2);
        const blob = new Blob([json], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function startNetworkMonitor() {
        installNetworkMonitor();
        scriptState.networkMonitor.active = true;
        updateDevNetworkStatus('监听中。请在官方成绩页面点击“详情”、切换标签或触发你想捕获的操作。');
        updateDevNetworkDisplay();
    }

    function stopNetworkMonitor() {
        scriptState.networkMonitor.active = false;
        updateDevNetworkStatus(`已停止监听，当前记录 ${scriptState.rawData.networkCaptures.length} 条。`);
        updateDevNetworkDisplay();
    }

    function installNetworkMonitor() {
        if (scriptState.networkMonitor.installed) return;

        const pageWindow = getPageWindow();
        patchPageFetchForMonitor(pageWindow);
        patchPageXHRForMonitor(pageWindow);
        scriptState.networkMonitor.installed = true;
    }

    function patchPageFetchForMonitor(pageWindow) {
        if (!pageWindow || typeof pageWindow.fetch !== 'function') return;

        scriptState.networkMonitor.originalFetch = pageWindow.fetch;
        const originalFetch = pageWindow.fetch.bind(pageWindow);

        pageWindow.fetch = async function(input, init = {}) {
            const requestInfo = normalizeFetchRequestInfo(input, init);
            const started = Date.now();

            try {
                const response = await originalFetch(input, init);
                captureFetchResponse(requestInfo, response, started);
                return response;
            } catch (err) {
                recordNetworkCapture({
                    transport: 'fetch',
                    method: requestInfo.method,
                    url: requestInfo.url,
                    requestBody: requestInfo.body,
                    status: null,
                    networkError: true,
                    error: String(err),
                    durationMs: Date.now() - started,
                    responseText: ''
                });
                throw err;
            }
        };
    }

    function patchPageXHRForMonitor(pageWindow) {
        if (!pageWindow || !pageWindow.XMLHttpRequest) return;

        const proto = pageWindow.XMLHttpRequest.prototype;
        if (proto.__szuScoreMonitorPatched) return;

        scriptState.networkMonitor.originalXHROpen = proto.open;
        scriptState.networkMonitor.originalXHRSend = proto.send;

        proto.open = function(method, url) {
            this.__szuScoreMonitorInfo = {
                method: method || 'GET',
                url: normalizeSameOriginUrl(url) || String(url || ''),
                started: null,
                requestBody: ''
            };
            return scriptState.networkMonitor.originalXHROpen.apply(this, arguments);
        };

        proto.send = function(body) {
            const info = this.__szuScoreMonitorInfo || {};
            info.started = Date.now();
            info.requestBody = typeof body === 'string' ? body : '';

            this.addEventListener('loadend', () => {
                const responseText = readXHRResponseText(this);
                recordNetworkCapture({
                    transport: 'XMLHttpRequest',
                    method: info.method || 'GET',
                    url: info.url || '',
                    requestBody: info.requestBody || '',
                    status: this.status || null,
                    networkError: false,
                    error: null,
                    durationMs: Date.now() - (info.started || Date.now()),
                    responseText
                });
            });

            return scriptState.networkMonitor.originalXHRSend.apply(this, arguments);
        };

        proto.__szuScoreMonitorPatched = true;
    }

    function normalizeFetchRequestInfo(input, init) {
        let url = '';
        let method = 'GET';
        let body = '';

        if (typeof input === 'string') {
            url = input;
        } else if (input && input.url) {
            url = input.url;
            method = input.method || method;
        }

        if (init && init.method) {
            method = init.method;
        }

        if (init && typeof init.body === 'string') {
            body = init.body;
        }

        return {
            url: normalizeSameOriginUrl(url) || String(url || ''),
            method,
            body
        };
    }

    async function captureFetchResponse(requestInfo, response, started) {
        if (!shouldCaptureNetworkUrl(requestInfo.url)) return;

        let responseText = '';
        try {
            const clone = response.clone();
            responseText = await clone.text();
        } catch (err) {
            responseText = `[response text unavailable: ${String(err)}]`;
        }

        recordNetworkCapture({
            transport: 'fetch',
            method: requestInfo.method,
            url: requestInfo.url,
            requestBody: requestInfo.body,
            status: response.status || null,
            networkError: false,
            error: null,
            durationMs: Date.now() - started,
            responseText
        });
    }

    function readXHRResponseText(xhr) {
        try {
            const responseType = xhr.responseType || '';
            if (responseType === '' || responseType === 'text') {
                return xhr.responseText || '';
            }
            return `[non-text responseType: ${responseType}]`;
        } catch (err) {
            return `[response text unavailable: ${String(err)}]`;
        }
    }

    function recordNetworkCapture(capture) {
        if (!scriptState.networkMonitor.active) return;
        if (!shouldCaptureNetworkUrl(capture.url)) return;

        const responseText = capture.responseText || '';
        const item = {
            timestamp: new Date().toISOString(),
            transport: capture.transport,
            method: capture.method,
            url: safeUrlForReport(capture.url),
            requestBody: capture.requestBody || '',
            status: capture.status,
            networkError: capture.networkError || false,
            error: capture.error || null,
            durationMs: capture.durationMs,
            responseTextLength: responseText.length,
            responseText,
            summary: summarizeCapturedResponse(responseText)
        };

        scriptState.rawData.networkCaptures.push(item);
        updateDevNetworkDisplay();
    }

    function shouldCaptureNetworkUrl(url) {
        try {
            const parsed = new URL(url, location.href);
            if (parsed.origin !== location.origin) return false;
            return parsed.pathname.includes('/sys/cjcx/');
        } catch (e) {
            return false;
        }
    }

    function summarizeCapturedResponse(responseText) {
        if (!responseText) {
            return { isJson: false, responseLength: 0, keys: [], rowCandidates: [] };
        }

        try {
            const data = JSON.parse(responseText);
            return {
                isJson: true,
                responseLength: responseText.length,
                keys: data && typeof data === 'object' ? Object.keys(data).slice(0, 30) : [],
                rowCandidates: collectRowCandidates(data).slice(0, 20),
                coefficientLikeFields: collectCoefficientLikeFields(data).slice(0, 80)
            };
        } catch (err) {
            return {
                isJson: false,
                responseLength: responseText.length,
                textPreview: maskSensitiveText(responseText.slice(0, 800))
            };
        }
    }

    function updateDevNetworkDisplay() {
        if (!scriptState.container) return;

        const dataEl = scriptState.container.querySelector('#dev-network-data');
        const countEl = scriptState.container.querySelector('#dev-network-count');
        const downloadBtn = scriptState.container.querySelector('#dev-download-network-captures');
        const startBtn = scriptState.container.querySelector('#dev-start-network-monitor');
        const stopBtn = scriptState.container.querySelector('#dev-stop-network-monitor');
        const captures = scriptState.rawData.networkCaptures;

        if (countEl) countEl.textContent = captures.length;
        if (downloadBtn) downloadBtn.disabled = captures.length === 0;
        if (startBtn) startBtn.disabled = scriptState.networkMonitor.active;
        if (stopBtn) stopBtn.disabled = !scriptState.networkMonitor.active;

        if (dataEl) {
            dataEl.textContent = captures.length
                ? JSON.stringify(captures.slice(-30), null, 2)
                : '暂无数据';
        }
    }

    function updateDevNetworkStatus(message) {
        if (!scriptState.container) return;
        const statusEl = scriptState.container.querySelector('#dev-network-status');
        if (statusEl) statusEl.textContent = message;
    }

    function buildNetworkCaptureExport() {
        return {
            timestamp: new Date().toISOString(),
            meta: {
                scriptVersion: '4.12-network-monitor',
                origin: location.origin,
                pathname: location.pathname,
                pageTitle: document.title,
                note: '页面请求监听结果；包含匹配 /sys/cjcx/ 的请求 body 和完整响应正文；不导出 Cookie、Token、Authorization。'
            },
            runtime: inspectPageRuntimeForOfficialProbe(),
            captures: scriptState.rawData.networkCaptures
        };
    }

    function dismissProbeNetworkErrorDialogs() {
        try {
            const candidates = Array.from(document.querySelectorAll('button, input[type="button"], [role="button"], .bh-btn, .jqx-button'));
            for (const element of candidates) {
                const text = (element.textContent || element.value || '').trim();
                if (!/^(关闭|确定|OK)$/i.test(text)) continue;

                if (hasNetworkErrorAncestor(element)) {
                    element.click();
                    return;
                }
            }

            const closeCandidates = Array.from(document.querySelectorAll('[class*="close"], [aria-label*="关闭"], [title*="关闭"]'));
            for (const element of closeCandidates) {
                if (hasNetworkErrorAncestor(element)) {
                    element.click();
                    return;
                }
            }
        } catch (err) {
            console.warn('[深大成绩查询] 自动关闭网络错误弹窗失败:', err);
        }
    }

    function hasNetworkErrorAncestor(element) {
        let current = element;
        let depth = 0;
        while (current && depth < 8) {
            const text = current.textContent || '';
            if (text.includes('网络错误')) return true;
            current = current.parentElement;
            depth++;
        }
        return false;
    }

    async function runCoefficientEndpointProbe(updateStatus, onProgress) {
        const startedAt = new Date();
        const status = typeof updateStatus === 'function' ? updateStatus : () => {};
        const notifyProgress = typeof onProgress === 'function' ? onProgress : () => {};
        const report = {
            timestamp: startedAt.toISOString(),
            completedAt: null,
            state: 'running',
            meta: {
                scriptVersion: '4.11-dev-probe',
                origin: location.origin,
                pathname: location.pathname,
                pageTitle: document.title,
                note: '结果用于定位新成绩系数接口；优先复刻官方 jxblrcjxs.do 和 BH_UTILS.doSyncAjax 调用；导出每个探测请求的完整响应正文 rawResponseText；不导出 Cookie、Token、Authorization。'
            },
            seedCourses: [],
            discovery: {
                resourcesScanned: [],
                discoveredEndpoints: []
            },
            officialCoefficientProbe: null,
            candidates: [],
            payloadTemplates: [],
            requestCount: 0,
            maxRequests: 80,
            promising: [],
            requests: []
        };
        notifyProgress(report);

        status('正在获取课程列表作为探测种子...');
        const initialCourses = await ensureInitialCoursesForProbe();
        const seedCourses = buildProbeCourseSeeds(initialCourses);
        const primarySeed = seedCourses[0] || {};
        report.seedCourses = seedCourses;
        notifyProgress(report);

        status('正在复刻官方教学班成绩系数接口调用...');
        report.officialCoefficientProbe = await runOfficialCoefficientProbe(seedCourses, status, (partialProbe) => {
            report.officialCoefficientProbe = partialProbe;
            notifyProgress(report);
        });
        notifyProgress(report);

        status('正在扫描当前页面和脚本资源，寻找候选接口...');
        const discovery = await discoverCoefficientEndpointCandidates(status, (partialDiscovery) => {
            report.discovery = partialDiscovery;
            notifyProgress(report);
        });
        report.discovery = discovery;

        const candidates = buildCoefficientProbeCandidates(discovery.discoveredEndpoints);
        const payloadTemplates = buildCoefficientProbePayloads(primarySeed);
        report.candidates = candidates.map(safeUrlForReport);
        report.payloadTemplates = payloadTemplates.map(payload => ({
            name: payload.name,
            keys: payload.keys,
            dataLength: payload.data.length
        }));
        notifyProgress(report);

        let requestCount = 0;

        for (const endpoint of candidates) {
            const payloads = selectProbePayloadsForEndpoint(endpoint, payloadTemplates);

            for (const payload of payloads) {
                if (requestCount >= report.maxRequests) break;

                requestCount++;
                status(`正在探测接口 ${requestCount}/${report.maxRequests}：${shortEndpointName(endpoint)} / ${payload.name}`);

                const response = await gmProbeRequest({
                    method: 'POST',
                    url: endpoint,
                    data: payload.data,
                    timeout: 6000
                });
                dismissProbeNetworkErrorDialogs();

                const summary = summarizeProbeHttpResponse(response);
                const result = {
                    index: requestCount,
                    endpoint: safeUrlForReport(endpoint),
                    method: 'POST',
                    payloadName: payload.name,
                    payloadKeys: payload.keys,
                    status: response.status || null,
                    durationMs: response.durationMs,
                    networkError: response.networkError || false,
                    error: response.error || null,
                    finalUrl: response.finalUrl ? safeUrlForReport(response.finalUrl) : null,
                    requestPayload: payload.data,
                    rawResponseTextLength: response.responseText ? response.responseText.length : 0,
                    rawResponseText: response.responseText || '',
                    summary
                };
                result.match = scoreProbeResult(result);
                report.requests.push(result);
                report.requestCount = requestCount;
                refreshProbePromising(report);
                notifyProgress(report);

                await sleep(80);
            }

            if (requestCount >= report.maxRequests) break;
        }

        report.state = 'completed';
        report.completedAt = new Date().toISOString();
        refreshProbePromising(report);
        notifyProgress(report);
        return report;
    }

    async function ensureInitialCoursesForProbe() {
        const cachedRows = scriptState.rawData.initialCourses?.datas?.xscjcx?.rows;
        if (Array.isArray(cachedRows) && cachedRows.length > 0) {
            return cachedRows;
        }

        const rows = await fetchInitialCourseList();
        return Array.isArray(rows) ? rows : [];
    }

    function buildProbeCourseSeeds(courses) {
        if (!Array.isArray(courses)) return [];

        const seedFields = [
            'JXBID', 'XNXQDM', 'XNXQDM_DISPLAY', 'KCH', 'KCDM', 'KCM',
            'KXH', 'KKDWDM', 'KCXZDM', 'KCLBDM', 'XF'
        ];

        return courses.slice(0, 3).map((course, index) => {
            const seed = { index };
            seedFields.forEach(field => {
                if (course && course[field] !== undefined && course[field] !== null) {
                    seed[field] = String(course[field]);
                }
            });
            return seed;
        });
    }

    async function runOfficialCoefficientProbe(seedCourses, updateStatus, onProgress) {
        const status = typeof updateStatus === 'function' ? updateStatus : () => {};
        const notifyProgress = typeof onProgress === 'function' ? onProgress : () => {};
        const courses = (seedCourses || [])
            .filter(course => course && course.JXBID)
            .slice(0, 5);
        const endpoints = buildOfficialCoefficientEndpointUrls();
        const transports = buildOfficialCoefficientTransports();
        const probe = {
            state: courses.length > 0 ? 'running' : 'skipped',
            reason: courses.length > 0 ? null : '没有可用于探测的 JXBID',
            endpointPurpose: '官方前端 getJxbLrcjxs({JXBID: jxbid, XSYC: 0})',
            expectedShape: 'datas.jxblrcjxs.rows[0]，行内以 XS 结尾的数字字段为成绩项系数',
            runtime: inspectPageRuntimeForOfficialProbe(),
            courseCount: courses.length,
            endpoints: endpoints.map(safeUrlForReport),
            transports: transports.map(transport => transport.name),
            requestCount: 0,
            successCount: 0,
            coefficientHits: [],
            requests: []
        };
        notifyProgress(probe);

        if (courses.length === 0) {
            return probe;
        }

        const totalRequests = courses.length * endpoints.length * transports.length;

        for (const course of courses) {
            const payloadData = { JXBID: course.JXBID, XSYC: '0' };
            const payload = encodeFormData(payloadData);

            for (const endpoint of endpoints) {
                for (const transport of transports) {
                    probe.requestCount++;
                    status(`官方系数接口专项探测 ${probe.requestCount}/${totalRequests}：${transport.name} / ${shortEndpointName(endpoint)} / ${course.KCM || course.JXBID}`);

                    const response = await transport.request({
                        method: 'POST',
                        url: endpoint,
                        data: payload,
                        dataObject: payloadData,
                        timeout: 10000
                    });
                    dismissProbeNetworkErrorDialogs();
                    const summary = summarizeProbeHttpResponse(response);
                    const parsed = parseOfficialCoefficientResponse(response);
                    const result = {
                        index: probe.requestCount,
                        transport: transport.name,
                        endpoint: safeUrlForReport(endpoint),
                        method: 'POST',
                        course: {
                            index: course.index,
                            JXBID: course.JXBID,
                            KCM: course.KCM || '',
                            KCH: course.KCH || course.KCDM || '',
                            XNXQDM: course.XNXQDM || ''
                        },
                        requestPayload: payload,
                        status: response.status || null,
                        durationMs: response.durationMs,
                        networkError: response.networkError || false,
                        error: response.error || null,
                        finalUrl: response.finalUrl ? safeUrlForReport(response.finalUrl) : null,
                        rawResponseTextLength: response.responseText ? response.responseText.length : 0,
                        rawResponseText: response.responseText || '',
                        summary,
                        parsed
                    };
                    result.match = scoreOfficialCoefficientProbeResult(result);
                    probe.requests.push(result);

                    if (parsed.coefficientFields.length > 0) {
                        probe.successCount++;
                        probe.coefficientHits.push({
                            index: result.index,
                            transport: result.transport,
                            endpoint: result.endpoint,
                            course: result.course,
                            coefficientFields: parsed.coefficientFields
                        });
                    }

                    notifyProgress(probe);
                    await sleep(120);
                }
            }
        }

        probe.state = 'completed';
        notifyProgress(probe);
        return probe;
    }

    function buildOfficialCoefficientEndpointUrls() {
        const urls = new Set();
        const appModulePath = `${location.origin}/jwapp/sys/cjcx/modules/cjcx/jxblrcjxs.do`;
        const defaultModulePath = `${location.origin}/jwapp/sys/cjcx/*default/modules/cjcx/jxblrcjxs.do`;
        urls.add(appModulePath);
        urls.add(defaultModulePath);

        try {
            const pageWindow = getPageWindow();
            const absPath = pageWindow?.WIS_EMAP_SERV?.getAbsPath?.('/modules/cjcx/jxblrcjxs.do');
            if (absPath) {
                urls.add(new URL(absPath, location.href).toString());
            }
        } catch (e) {
            console.warn('[深大成绩查询] 读取 WIS_EMAP_SERV.getAbsPath 失败:', e);
        }

        try {
            const pageWindow = getPageWindow();
            const modulePath = pageWindow?.APP_CONFIG?.MODULE_PATH;
            if (modulePath) {
                urls.add(new URL('cjcx/jxblrcjxs.do', modulePath).toString());
            }
        } catch (e) {
            console.warn('[深大成绩查询] 读取 APP_CONFIG.MODULE_PATH 失败:', e);
        }

        return Array.from(urls)
            .map(url => normalizeSameOriginUrl(url))
            .filter(Boolean)
            .filter((url, index, arr) => arr.indexOf(url) === index);
    }

    function buildOfficialCoefficientTransports() {
        return [
            { name: 'page-BH_UTILS-doSyncAjax', request: bhDoSyncAjaxProbeRequest },
            { name: 'GM_xmlhttpRequest', request: gmProbeRequest },
            { name: 'page-fetch', request: pageFetchProbeRequest },
            { name: 'page-jquery-ajax', request: jqueryAjaxProbeRequest }
        ];
    }

    function getPageWindow() {
        if (typeof unsafeWindow !== 'undefined') {
            return unsafeWindow;
        }
        return window;
    }

    function inspectPageRuntimeForOfficialProbe() {
        const info = {
            hasUnsafeWindow: typeof unsafeWindow !== 'undefined',
            hasBH_UTILS: false,
            hasBHDoSyncAjax: false,
            hasWIS_EMAP_SERV: false,
            hasWISGetAbsPath: false,
            wisAbsPathSample: null,
            hasAPP_CONFIG: false,
            appModulePath: null,
            hasJQuery: false,
            hasFetch: false,
            hasRequire: false
        };

        try {
            const pageWindow = getPageWindow();
            info.hasBH_UTILS = !!pageWindow?.BH_UTILS;
            info.hasBHDoSyncAjax = typeof pageWindow?.BH_UTILS?.doSyncAjax === 'function';
            info.hasWIS_EMAP_SERV = !!pageWindow?.WIS_EMAP_SERV;
            info.hasWISGetAbsPath = typeof pageWindow?.WIS_EMAP_SERV?.getAbsPath === 'function';
            if (info.hasWISGetAbsPath) {
                info.wisAbsPathSample = String(pageWindow.WIS_EMAP_SERV.getAbsPath('/modules/cjcx/jxblrcjxs.do'));
            }
            info.hasAPP_CONFIG = !!pageWindow?.APP_CONFIG;
            info.appModulePath = pageWindow?.APP_CONFIG?.MODULE_PATH ? String(pageWindow.APP_CONFIG.MODULE_PATH) : null;
            info.hasJQuery = !!(pageWindow?.jQuery || pageWindow?.$);
            info.hasFetch = typeof pageWindow?.fetch === 'function';
            info.hasRequire = typeof pageWindow?.require === 'function';
        } catch (err) {
            info.error = String(err);
        }

        return info;
    }

    function bhDoSyncAjaxProbeRequest(options) {
        const started = Date.now();

        try {
            const pageWindow = getPageWindow();
            const bhUtils = pageWindow?.BH_UTILS;
            if (!bhUtils || typeof bhUtils.doSyncAjax !== 'function') {
                return Promise.resolve({
                    networkError: true,
                    error: 'BH_UTILS.doSyncAjax unavailable',
                    responseText: '',
                    durationMs: Date.now() - started
                });
            }

            const params = options.dataObject || decodeFormData(options.data || '');
            const response = bhUtils.doSyncAjax(options.url, params);
            const responseText = response === undefined
                ? ''
                : typeof response === 'string'
                ? response
                : JSON.stringify(response);

            return Promise.resolve({
                status: 200,
                responseText,
                responseHeaders: '',
                finalUrl: options.url,
                durationMs: Date.now() - started,
                syntheticTransport: true
            });
        } catch (err) {
            return Promise.resolve({
                networkError: true,
                error: String(err),
                responseText: '',
                durationMs: Date.now() - started
            });
        }
    }

    async function pageFetchProbeRequest(options) {
        const started = Date.now();
        const method = options.method || 'GET';
        const headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'X-Requested-With': 'XMLHttpRequest'
        };

        if (method.toUpperCase() === 'POST') {
            headers['Content-Type'] = 'application/x-www-form-urlencoded;charset=UTF-8';
        }

        try {
            const pageWindow = getPageWindow();
            const fetchImpl = pageWindow.fetch ? pageWindow.fetch.bind(pageWindow) : fetch.bind(window);
            const controller = new AbortController();
            const timer = setTimeout(() => controller.abort(), options.timeout || 10000);
            const response = await fetchImpl(options.url, {
                method,
                headers,
                body: options.data || undefined,
                credentials: 'same-origin',
                cache: 'no-store',
                signal: controller.signal
            });
            clearTimeout(timer);

            const responseText = await response.text();
            const responseHeaders = [];
            response.headers.forEach((value, key) => {
                responseHeaders.push(`${key}: ${value}`);
            });

            return {
                status: response.status,
                responseText,
                responseHeaders: responseHeaders.join('\n'),
                finalUrl: response.url || options.url,
                durationMs: Date.now() - started
            };
        } catch (err) {
            return {
                networkError: true,
                error: err && err.name === 'AbortError' ? 'timeout' : String(err),
                responseText: '',
                durationMs: Date.now() - started
            };
        }
    }

    function jqueryAjaxProbeRequest(options) {
        const started = Date.now();
        const method = options.method || 'GET';
        const headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'X-Requested-With': 'XMLHttpRequest'
        };

        return new Promise(resolve => {
            try {
                const pageWindow = getPageWindow();
                const jq = pageWindow.jQuery || pageWindow.$;
                if (!jq || typeof jq.ajax !== 'function') {
                    resolve({
                        networkError: true,
                        error: 'jQuery ajax unavailable',
                        responseText: '',
                        durationMs: Date.now() - started
                    });
                    return;
                }

                jq.ajax({
                    url: options.url,
                    type: method,
                    method,
                    data: options.data || undefined,
                    dataType: 'text',
                    contentType: method.toUpperCase() === 'POST'
                        ? 'application/x-www-form-urlencoded;charset=UTF-8'
                        : undefined,
                    headers,
                    timeout: options.timeout || 10000,
                    xhrFields: { withCredentials: true }
                }).done((data, textStatus, jqXHR) => {
                    resolve({
                        status: jqXHR?.status || 200,
                        responseText: typeof data === 'string' ? data : JSON.stringify(data),
                        responseHeaders: jqXHR?.getAllResponseHeaders?.() || '',
                        finalUrl: options.url,
                        durationMs: Date.now() - started
                    });
                }).fail((jqXHR, textStatus, errorThrown) => {
                    resolve({
                        status: jqXHR?.status || null,
                        networkError: true,
                        error: errorThrown || textStatus || 'ajax failed',
                        responseText: jqXHR?.responseText || '',
                        responseHeaders: jqXHR?.getAllResponseHeaders?.() || '',
                        finalUrl: options.url,
                        durationMs: Date.now() - started
                    });
                });
            } catch (err) {
                resolve({
                    networkError: true,
                    error: String(err),
                    responseText: '',
                    durationMs: Date.now() - started
                });
            }
        });
    }

    function parseOfficialCoefficientResponse(response) {
        const result = {
            isJson: false,
            isOfficialShape: false,
            extCode: null,
            extMsg: null,
            rowCount: 0,
            rowKeys: [],
            coefficientFields: []
        };

        try {
            const data = JSON.parse(response.responseText || '');
            result.isJson = true;
            const table = data?.datas?.jxblrcjxs;
            if (!table) return result;

            result.isOfficialShape = true;
            result.extCode = table.extParams?.code ?? null;
            result.extMsg = table.extParams?.msg ?? null;
            const rows = Array.isArray(table.rows) ? table.rows : [];
            result.rowCount = rows.length;
            result.rowKeys = rows[0] ? Object.keys(rows[0]) : [];
            result.coefficientFields = rows.flatMap((row, rowIndex) =>
                extractOfficialCoefficientFields(row).map(field => ({
                    rowIndex,
                    ...field
                }))
            );
        } catch (e) {
            result.parseError = e.message;
        }

        return result;
    }

    function extractOfficialCoefficientFields(row) {
        if (!row || typeof row !== 'object') return [];

        return Object.keys(row)
            .filter(key => /^(PSCJ|QZCJ|QMCJ|SYCJ|SJCJ|QTCJ\d+)XS$/i.test(key))
            .filter(key => row[key] !== null && row[key] !== undefined && row[key] !== '' && !isNaN(parseFloat(row[key])))
            .map(key => ({
                key,
                scoreItem: key.replace(/XS$/i, ''),
                value: parseFloat(row[key])
            }));
    }

    function scoreOfficialCoefficientProbeResult(result) {
        const reasons = [];
        let score = 0;

        if (result.status === 200) {
            score += 1;
            reasons.push('HTTP 200');
        }

        if (result.parsed.isOfficialShape) {
            score += 3;
            reasons.push('官方 jxblrcjxs 结构');
        }

        if (result.parsed.rowCount > 0) {
            score += 2;
            reasons.push('返回教学班系数行');
        }

        if (result.parsed.coefficientFields.length > 0) {
            score += 8;
            reasons.push('发现官方 *XS 系数字段');
        }

        return { score, reasons };
    }

    async function discoverCoefficientEndpointCandidates(updateStatus, onProgress) {
        const resources = collectProbeResourceUrls();
        const discovered = new Set();
        const scans = [];
        const notifyProgress = typeof onProgress === 'function' ? onProgress : () => {};

        for (let i = 0; i < resources.length; i++) {
            const resourceUrl = resources[i];
            updateStatus(`正在扫描资源 ${i + 1}/${resources.length}：${shortEndpointName(resourceUrl)}`);

            const response = await gmProbeRequest({
                method: 'GET',
                url: resourceUrl,
                timeout: 6000
            });

            const endpoints = response.responseText
                ? extractEndpointCandidates(response.responseText).slice(0, 80)
                : [];

            endpoints.forEach(endpoint => discovered.add(endpoint));
            scans.push({
                url: safeUrlForReport(resourceUrl),
                status: response.status || null,
                durationMs: response.durationMs,
                responseLength: response.responseText ? response.responseText.length : 0,
                endpointsFound: endpoints.map(safeUrlForReport),
                rawResponseText: response.responseText || ''
            });
            notifyProgress({
                resourcesScanned: scans,
                discoveredEndpoints: Array.from(discovered)
            });

            await sleep(60);
        }

        return {
            resourcesScanned: scans,
            discoveredEndpoints: Array.from(discovered)
        };
    }

    function refreshProbePromising(report) {
        report.promising = report.requests
            .filter(item => item.match && item.match.score > 0)
            .sort((a, b) => b.match.score - a.match.score)
            .slice(0, 20);
    }

    function collectProbeResourceUrls() {
        const urls = new Set();
        urls.add(location.origin + location.pathname);

        Array.from(document.scripts || []).forEach(script => {
            if (script.src) urls.add(script.src);
        });

        if (window.performance && typeof window.performance.getEntriesByType === 'function') {
            performance.getEntriesByType('resource').forEach(entry => {
                if (entry && entry.name) urls.add(entry.name);
            });
        }

        return Array.from(urls)
            .map(url => normalizeSameOriginUrl(url))
            .filter(Boolean)
            .filter(url => {
                try {
                    const parsed = new URL(url);
                    const path = parsed.pathname.toLowerCase();
                    return path.includes('/jwapp/sys/cjcx') || path.endsWith('.js');
                } catch (e) {
                    return false;
                }
            })
            .filter((url, index, arr) => arr.indexOf(url) === index)
            .slice(0, 18);
    }

    function extractEndpointCandidates(text) {
        const endpoints = new Set();
        const source = String(text || '')
            .replace(/\\u002F/g, '/')
            .replace(/\\\//g, '/')
            .slice(0, 300000);

        const patterns = [
            /\/jwapp\/sys\/cjcx\/[A-Za-z0-9_./-]+\.do/g,
            /modules\/cjcx\/[A-Za-z0-9_./-]+\.do/g,
            /["'`]([A-Za-z0-9_./-]*(?:cjcx|xscj|jxbl|jxb|cjxs|cjbl)[A-Za-z0-9_./-]*\.do)["'`]/gi
        ];

        patterns.forEach(pattern => {
            let match;
            while ((match = pattern.exec(source)) !== null) {
                const raw = match[1] || match[0];
                const normalized = normalizeEndpointCandidate(raw);
                if (normalized) endpoints.add(normalized);
            }
        });

        return Array.from(endpoints);
    }

    function buildCoefficientProbeCandidates(discoveredEndpoints) {
        const hardcoded = getDefaultCoefficientProbeEndpoints();
        const ordered = [
            ...Array.from(discoveredEndpoints || []),
            ...hardcoded
        ];

        return ordered
            .map(endpoint => normalizeEndpointCandidate(endpoint))
            .filter(Boolean)
            .filter((endpoint, index, arr) => arr.indexOf(endpoint) === index)
            .slice(0, 30);
    }

    function getDefaultCoefficientProbeEndpoints() {
        const base = `${location.origin}/jwapp/sys/cjcx/modules/cjcx/`;
        return [
            'jxblrcjxs.do',
            'xscjcx.do',
            'xscjcxmx.do',
            'xscjmx.do',
            'cjcxmx.do',
            'cjmx.do',
            'cjjg.do',
            'cjxx.do',
            'cjxs.do',
            'cjxsxx.do',
            'xscjxs.do',
            'xscjbl.do',
            'cjbl.do',
            'jxcjxs.do',
            'jxblxs.do',
            'jxblxx.do',
            'jxblrcjxsck.do',
            'kcjxcjxs.do',
            'xskccjmx.do',
            'kccjmx.do',
            'jxbxx.do'
        ].map(name => base + name);
    }

    function buildCoefficientProbePayloads(seed) {
        const payloads = [];
        const jxbid = seed?.JXBID || '';
        const xnxqdm = seed?.XNXQDM || '';
        const kch = seed?.KCH || seed?.KCDM || '';

        if (jxbid) {
            payloads.push({
                name: 'jxbid-xsyc',
                keys: ['JXBID', 'XSYC'],
                data: encodeFormData({ JXBID: jxbid, XSYC: '0' })
            });
            payloads.push({
                name: 'jxbid-only',
                keys: ['JXBID'],
                data: encodeFormData({ JXBID: jxbid })
            });
            payloads.push({
                name: 'query-jxbid',
                keys: ['querySetting', 'pageSize', 'pageNumber'],
                data: buildQuerySettingPayload([{ name: 'JXBID', value: jxbid, linkOpt: 'and', builder: 'equal' }])
            });
        }

        if (jxbid && xnxqdm) {
            payloads.push({
                name: 'query-jxbid-xnxq',
                keys: ['querySetting', 'pageSize', 'pageNumber'],
                data: buildQuerySettingPayload([
                    { name: 'JXBID', value: jxbid, linkOpt: 'and', builder: 'equal' },
                    { name: 'XNXQDM', value: xnxqdm, linkOpt: 'and', builder: 'equal' }
                ])
            });
            payloads.push({
                name: 'jxbid-xnxqdm',
                keys: ['JXBID', 'XNXQDM'],
                data: encodeFormData({ JXBID: jxbid, XNXQDM: xnxqdm })
            });
        }

        if (kch && xnxqdm) {
            payloads.push({
                name: 'query-kch-xnxq',
                keys: ['querySetting', 'pageSize', 'pageNumber'],
                data: buildQuerySettingPayload([
                    { name: 'KCH', value: kch, linkOpt: 'and', builder: 'equal' },
                    { name: 'XNXQDM', value: xnxqdm, linkOpt: 'and', builder: 'equal' }
                ])
            });
        }

        payloads.push({
            name: 'page-list',
            keys: ['pageSize', 'pageNumber'],
            data: encodeFormData({ pageSize: '20', pageNumber: '1' })
        });

        return payloads;
    }

    function selectProbePayloadsForEndpoint(endpoint, payloadTemplates) {
        if (!payloadTemplates || payloadTemplates.length === 0) return [];

        const lowerEndpoint = endpoint.toLowerCase();
        const preferredNames = lowerEndpoint.includes('xscjcx.do')
            ? ['page-list', 'query-jxbid', 'query-jxbid-xnxq', 'query-kch-xnxq']
            : ['jxbid-xsyc', 'jxbid-only', 'query-jxbid', 'query-jxbid-xnxq'];

        const selected = preferredNames
            .map(name => payloadTemplates.find(payload => payload.name === name))
            .filter(Boolean);

        if (selected.length > 0) {
            return selected.slice(0, 4);
        }

        return payloadTemplates.slice(0, 3);
    }

    function buildQuerySettingPayload(settings) {
        return encodeFormData({
            querySetting: JSON.stringify(settings),
            pageSize: '20',
            pageNumber: '1'
        });
    }

    function encodeFormData(data) {
        return Object.keys(data)
            .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(data[key])}`)
            .join('&');
    }

    function decodeFormData(formText) {
        const result = {};
        String(formText || '').split('&').forEach(pair => {
            if (!pair) return;
            const parts = pair.split('=');
            const key = decodeURIComponent(parts[0] || '');
            const value = decodeURIComponent(parts.slice(1).join('=') || '');
            if (key) result[key] = value;
        });
        return result;
    }

    function gmProbeRequest(options) {
        const started = Date.now();
        const method = options.method || 'GET';
        const headers = {
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'X-Requested-With': 'XMLHttpRequest'
        };

        if (method.toUpperCase() === 'POST') {
            headers['Content-Type'] = 'application/x-www-form-urlencoded;charset=UTF-8';
        }

        return new Promise(resolve => {
            GM_xmlhttpRequest({
                method,
                url: options.url,
                headers,
                data: options.data || undefined,
                anonymous: false,
                timeout: options.timeout || 6000,
                onload: res => {
                    resolve({
                        status: res.status,
                        responseText: res.responseText || '',
                        responseHeaders: res.responseHeaders || '',
                        finalUrl: res.finalUrl || options.url,
                        durationMs: Date.now() - started
                    });
                },
                onerror: err => {
                    resolve({
                        networkError: true,
                        error: String(err),
                        responseText: '',
                        durationMs: Date.now() - started
                    });
                },
                ontimeout: () => {
                    resolve({
                        networkError: true,
                        error: 'timeout',
                        responseText: '',
                        durationMs: Date.now() - started
                    });
                }
            });
        });
    }

    function summarizeProbeHttpResponse(response) {
        const text = response.responseText || '';
        const summary = {
            responseLength: text.length,
            contentType: getHeaderValue(response.responseHeaders, 'content-type'),
            isJson: false,
            jsonShape: null,
            rowCandidates: [],
            coefficientLikeFields: [],
            textPreview: null
        };

        if (!text) return summary;

        try {
            const parsed = JSON.parse(text);
            summary.isJson = true;
            summary.jsonShape = buildJsonShape(parsed, 0, 4);
            summary.rowCandidates = collectRowCandidates(parsed).slice(0, 20);
            summary.coefficientLikeFields = collectCoefficientLikeFields(parsed).slice(0, 80);
        } catch (e) {
            summary.textPreview = maskSensitiveText(text.slice(0, 800));
        }

        return summary;
    }

    function buildJsonShape(value, depth, maxDepth) {
        const type = getValueType(value);

        if (depth >= maxDepth || value === null || type !== 'object') {
            if (Array.isArray(value)) {
                return {
                    type: 'array',
                    length: value.length,
                    sample: value.length > 0 ? buildJsonShape(value[0], depth + 1, maxDepth) : null
                };
            }
            return { type };
        }

        if (Array.isArray(value)) {
            return {
                type: 'array',
                length: value.length,
                sample: value.length > 0 ? buildJsonShape(value[0], depth + 1, maxDepth) : null
            };
        }

        const keys = Object.keys(value);
        const shape = {
            type: 'object',
            keys: keys.slice(0, 40),
            children: {}
        };

        keys.slice(0, 20).forEach(key => {
            shape.children[key] = buildJsonShape(value[key], depth + 1, maxDepth);
        });

        if (keys.length > 20) {
            shape.omittedKeyCount = keys.length - 20;
        }

        return shape;
    }

    function collectRowCandidates(value, path = '$', results = [], depth = 0) {
        if (results.length >= 50 || depth > 8 || value === null || value === undefined) {
            return results;
        }

        if (Array.isArray(value)) {
            const firstObject = value.find(item => item && typeof item === 'object' && !Array.isArray(item));
            if (firstObject) {
                const sampleKeys = Object.keys(firstObject);
                results.push({
                    path,
                    length: value.length,
                    sampleKeys,
                    coefficientLikeFields: extractCoefficientFieldsFromObject(firstObject, path)
                });
            }

            value.slice(0, 2).forEach((item, index) => {
                collectRowCandidates(item, `${path}[${index}]`, results, depth + 1);
            });
            return results;
        }

        if (typeof value === 'object') {
            Object.keys(value).slice(0, 30).forEach(key => {
                collectRowCandidates(value[key], `${path}.${key}`, results, depth + 1);
            });
        }

        return results;
    }

    function collectCoefficientLikeFields(value, path = '$', results = [], depth = 0) {
        if (results.length >= 120 || depth > 8 || value === null || value === undefined) {
            return results;
        }

        if (Array.isArray(value)) {
            value.slice(0, 3).forEach((item, index) => {
                collectCoefficientLikeFields(item, `${path}[${index}]`, results, depth + 1);
            });
            return results;
        }

        if (typeof value === 'object') {
            Object.keys(value).slice(0, 50).forEach(key => {
                const childPath = `${path}.${key}`;
                const childValue = value[key];
                if (isCoefficientLikeKey(key) && isPrimitiveValue(childValue)) {
                    results.push({
                        path: childPath,
                        key,
                        value: maskSensitiveValue(key, childValue)
                    });
                }
                collectCoefficientLikeFields(childValue, childPath, results, depth + 1);
            });
        }

        return results;
    }

    function extractCoefficientFieldsFromObject(obj, basePath) {
        if (!obj || typeof obj !== 'object') return [];

        return Object.keys(obj)
            .filter(key => isCoefficientLikeKey(key) && isPrimitiveValue(obj[key]))
            .map(key => ({
                path: `${basePath}[].${key}`,
                key,
                value: maskSensitiveValue(key, obj[key])
            }));
    }

    function scoreProbeResult(result) {
        const reasons = [];
        let score = 0;

        if (result.status === 200) {
            score += 1;
            reasons.push('HTTP 200');
        }

        if (result.summary.isJson) {
            score += 1;
            reasons.push('JSON响应');
        }

        if (result.summary.rowCandidates.length > 0) {
            score += 1;
            reasons.push('包含数组行结构');
        }

        if (result.summary.coefficientLikeFields.length > 0) {
            score += 5;
            reasons.push('发现疑似系数字段');
        }

        if (result.summary.rowCandidates.some(row => row.coefficientLikeFields.length > 0)) {
            score += 3;
            reasons.push('行结构内发现疑似系数字段');
        }

        return { score, reasons };
    }

    function isCoefficientLikeKey(key) {
        return /^(PSCJ|QZCJ|QMCJ|SYCJ|SJCJ|QTCJ\d+)XS$/i.test(key)
            || /(PSCJXS|QMCJXS|PSCJBL|QMCJBL|CJXS|CJBL|CJQZ|QZ|BL|BILI|BILV|比例|系数|WEIGHT|RATE|PERCENT)/i.test(key);
    }

    function isSensitiveKey(key) {
        return /^(XH|XM|SFZH|ZJHM|SJHM|LXDH|PHONE|MOBILE|TEL|EMAIL|COOKIE|SESSION|TOKEN|TICKET|AUTH|PASSWORD|PASS|SECRET|YHM|USERNAME|USERCODE)$/i.test(key);
    }

    function isPrimitiveValue(value) {
        return value === null || ['string', 'number', 'boolean'].includes(typeof value);
    }

    function maskSensitiveValue(key, value) {
        if (isSensitiveKey(key)) return '[REDACTED]';
        if (typeof value === 'string' && value.length > 120) {
            return value.slice(0, 120) + '...[truncated]';
        }
        return value;
    }

    function maskSensitiveText(text) {
        return String(text || '')
            .replace(/(token|ticket|session|password|authorization|cookie)=([^&\s"']+)/ig, '$1=[REDACTED]')
            .replace(/("?(?:XH|XM|SFZH|ZJHM|SJHM|TOKEN|TICKET|SESSION|PASSWORD|AUTH)"?\s*[:=]\s*)("[^"]+"|'[^']+'|[^,\s}]+)/ig, '$1[REDACTED]');
    }

    function getValueType(value) {
        if (Array.isArray(value)) return 'array';
        if (value === null) return 'null';
        return typeof value;
    }

    function getHeaderValue(headers, name) {
        if (!headers) return null;

        const target = name.toLowerCase();
        const line = String(headers)
            .split(/\r?\n/)
            .find(item => item.toLowerCase().startsWith(target + ':'));

        return line ? line.slice(line.indexOf(':') + 1).trim() : null;
    }

    function normalizeSameOriginUrl(rawUrl) {
        try {
            const parsed = new URL(rawUrl, location.href);
            if (parsed.origin !== location.origin) return null;
            parsed.hash = '';
            parsed.search = '';
            return parsed.toString();
        } catch (e) {
            return null;
        }
    }

    function normalizeEndpointCandidate(rawEndpoint) {
        if (!rawEndpoint) return null;

        let endpoint = String(rawEndpoint)
            .replace(/\\u002F/g, '/')
            .replace(/\\\//g, '/')
            .split('?')[0]
            .split('#')[0]
            .trim();

        if (!endpoint || !endpoint.endsWith('.do')) return null;

        if (endpoint.includes('modules/cjcx/')) {
            endpoint = endpoint.slice(endpoint.indexOf('modules/cjcx/'));
        }

        let url;
        if (/^https?:\/\//i.test(endpoint)) {
            url = endpoint;
        } else if (endpoint.startsWith('/')) {
            url = location.origin + endpoint;
        } else if (endpoint.startsWith('modules/cjcx/')) {
            url = `${location.origin}/jwapp/sys/cjcx/${endpoint}`;
        } else {
            const filename = endpoint.split('/').pop();
            url = `${location.origin}/jwapp/sys/cjcx/modules/cjcx/${filename}`;
        }

        try {
            const parsed = new URL(url);
            if (parsed.origin !== location.origin) return null;
            if (!parsed.pathname.includes('/jwapp/sys/cjcx/')) return null;
            return parsed.origin + parsed.pathname;
        } catch (e) {
            return null;
        }
    }

    function safeUrlForReport(url) {
        try {
            const parsed = new URL(url, location.href);
            return parsed.origin + parsed.pathname;
        } catch (e) {
            return String(url || '');
        }
    }

    function shortEndpointName(url) {
        try {
            const parsed = new URL(url, location.href);
            const parts = parsed.pathname.split('/').filter(Boolean);
            return parts.slice(-2).join('/');
        } catch (e) {
            return String(url || '').slice(-60);
        }
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    toggleBtn.addEventListener('click', () => scriptState.container.classList.toggle('hidden'));

    /**
     * 根据平时成绩、期末成绩和总成绩推断系数
     * @param {number} pscj 平时成绩
     * @param {number} qmcj 期末成绩
     * @param {number} zcj 总成绩
     * @returns {object|null} 推断的系数 {pscjxs, qmcjxs} 或 null（无法推断）
     */
    function inferCoefficients(pscj, qmcj, zcj) {
        // 常见的系数比例（平时:期末）
        const commonRatios = [
            { pscjxs: 10, qmcjxs: 90 },
            { pscjxs: 20, qmcjxs: 80 },
            { pscjxs: 30, qmcjxs: 70 },
            { pscjxs: 40, qmcjxs: 60 },
            { pscjxs: 50, qmcjxs: 50 },
            { pscjxs: 60, qmcjxs: 40 },
            { pscjxs: 70, qmcjxs: 30 },
            { pscjxs: 80, qmcjxs: 20 },
            { pscjxs: 90, qmcjxs: 10 },
            { pscjxs: 100, qmcjxs: 0 },
            { pscjxs: 0, qmcjxs: 100 }
        ];
        
        // 计算加权平均并四舍五入
        function calculateWeightedScore(p, q, pxs, qxs) {
            return Math.round((p * pxs / 100) + (q * qxs / 100));
        }
        
        // 1. 首先尝试常见比例
        for (const ratio of commonRatios) {
            const calculated = calculateWeightedScore(pscj, qmcj, ratio.pscjxs, ratio.qmcjxs);
            if (calculated === zcj) {
                console.log(`[系数推断] 匹配常见比例 ${ratio.pscjxs}:${ratio.qmcjxs}, 计算=${calculated}, 总成绩=${zcj}`);
                return ratio;
            }
        }
        
        // 2. 如果常见比例都不匹配，逐个尝试从1到99的平时成绩系数
        for (let pxs = 1; pxs <= 99; pxs++) {
            const qxs = 100 - pxs;
            const calculated = calculateWeightedScore(pscj, qmcj, pxs, qxs);
            if (calculated === zcj) {
                console.log(`[系数推断] 匹配比例 ${pxs}:${qxs}, 计算=${calculated}, 总成绩=${zcj}`);
                return { pscjxs: pxs, qmcjxs: qxs };
            }
        }
        
        // 3. 检查是否只有一种成绩（100%比例的情况）
        if (Math.round(pscj) === zcj) {
            console.log(`[系数推断] 可能是100%平时成绩`);
            return { pscjxs: 100, qmcjxs: 0 };
        }
        if (Math.round(qmcj) === zcj) {
            console.log(`[系数推断] 可能是100%期末成绩`);
            return { pscjxs: 0, qmcjxs: 100 };
        }
        
        // 无法推断
        console.log(`[系数推断] 无法推断系数: 平时=${pscj}, 期末=${qmcj}, 总成绩=${zcj}`);
        return null;
    }

    // 获取初始课程列表
    function fetchInitialCourseList() {
        return new Promise((resolve, reject) => {
            const url = `${location.origin}/jwapp/sys/cjcx/modules/cjcx/xscjcx.do`;
            console.log('[深大成绩查询] 正在获取初始课程列表:', url);
            
            GM_xmlhttpRequest({
                method: "POST",
                url: url,
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "X-Requested-With": "XMLHttpRequest"
                },
                data: "pageSize=100&pageNumber=1",
                timeout: 30000,
                onload: res => {
                    console.log('[深大成绩查询] 初始课程列表响应状态:', res.status);
                    try {
                        if (res.status !== 200) {
                            console.error('[深大成绩查询] 请求返回非200状态:', res.status, res.responseText);
                            reject(new Error(`请求失败，状态码: ${res.status}`));
                            return;
                        }
                        const data = JSON.parse(res.responseText);
                        console.log('[深大成绩查询] 解析成功，课程数量:', data?.datas?.xscjcx?.rows?.length || 0);
                        scriptState.rawData.initialCourses = data;
                        if (scriptState.devMode) {
                            updateDevDataDisplay();
                        }
                        resolve(data?.datas?.xscjcx?.rows || []);
                    } catch (e) {
                        console.error('[深大成绩查询] 解析初始课程列表失败:', e, res.responseText?.substring(0, 500));
                        reject(new Error("解析初始课程列表失败: " + e.message));
                    }
                },
                onerror: (err) => {
                    console.error('[深大成绩查询] 获取初始课程列表网络错误:', err);
                    reject(new Error("获取初始课程列表网络请求失败"));
                },
                ontimeout: () => {
                    console.error('[深大成绩查询] 获取初始课程列表超时');
                    reject(new Error("获取初始课程列表请求超时"));
                }
            });
        });
    }

    // 获取单门课程的系数
    function fetchCourseCoefficients(jxbid) {
        return new Promise(resolve => {
            const url = `${location.origin}/jwapp/sys/cjcx/modules/cjcx/jxblrcjxs.do`;
            GM_xmlhttpRequest({
                method: "POST",
                url: url,
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "X-Requested-With": "XMLHttpRequest"
                },
                data: `JXBID=${jxbid}&XSYC=0`,
                timeout: 5000, // 系数查询超时时间短一点
                onload: res => {
                    try {
                        if (res.status === 200) {
                            const data = JSON.parse(res.responseText);
                            // 结构: datas.jxblrcjxs.rows[0]
                            const row = data?.datas?.jxblrcjxs?.rows?.[0];
                            if (row) {
                                resolve({
                                    pscjxs: row.PSCJXS,
                                    qmcjxs: row.QMCJXS
                                });
                                return;
                            }
                        }
                    } catch (e) {
                        console.error(`[深大成绩查询] 获取课程系数失败 JXBID=${jxbid}`, e);
                    }
                    resolve(null);
                },
                onerror: () => resolve(null),
                ontimeout: () => resolve(null)
            });
        });
    }

    // 执行成绩查询
    function performQuery(score, scoreType) {
        return new Promise(resolve => {
            const payload = `querySetting=[{"name":"${scoreType}","value":"${score}","linkOpt":"and","builder":"equal"}]&pageSize=100&pageNumber=1`;
            const url = `${location.origin}/jwapp/sys/cjcx/modules/cjcx/xscjcx.do`;
            
            GM_xmlhttpRequest({
                method: "POST",
                url: url,
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "X-Requested-With": "XMLHttpRequest"
                },
                data: payload,
                timeout: 15000,
                onload: res => {
                    try {
                        if (res.status !== 200) {
                            console.error(`[深大成绩查询] 查询${scoreType}=${score}返回非200:`, res.status);
                            if (scriptState.devMode) {
                                addQueryResult(score, scoreType, [], { error: `HTTP ${res.status}`, rawText: res.responseText });
                            }
                            resolve([]);
                            return;
                        }
                        const data = JSON.parse(res.responseText);
                        const rows = data?.datas?.xscjcx?.rows || [];
                        
                        // 开发者模式：记录查询结果
                        if (scriptState.devMode) {
                            addQueryResult(score, scoreType, rows, data);
                        }
                        
                        resolve(rows);
                    } catch (e) {
                        console.error(`解析${scoreType}=${score}的响应失败:`, e);
                        // 开发者模式：记录错误
                        if (scriptState.devMode) {
                            addQueryResult(score, scoreType, [], { error: e.message, rawText: res.responseText?.substring(0, 500) });
                        }
                        resolve([]);
                    }
                },
                onerror: (err) => {
                    console.error(`查询${scoreType}=${score}时网络请求失败:`, err);
                    // 开发者模式：记录网络错误
                    if (scriptState.devMode) {
                        addQueryResult(score, scoreType, [], { networkError: true, error: String(err) });
                    }
                    resolve([]);
                },
                ontimeout: () => {
                    console.error(`查询${scoreType}=${score}超时`);
                    if (scriptState.devMode) {
                        addQueryResult(score, scoreType, [], { timeout: true });
                    }
                    resolve([]);
                }
            });
        });
    }

    initContainer();
    installInlineScoreTab();
    
    // 注册菜单命令
    GM_registerMenuCommand("打开深大成绩查询", () => {
        if (scriptState.container) {
            scriptState.container.classList.remove('hidden');
        }
    });
    
    // 注册开发者模式菜单命令
    GM_registerMenuCommand("🔧 开启开发者模式", () => {
        if (scriptState.container) {
            const devToggleContainer = scriptState.container.querySelector('#dev-toggle-container');
            if (devToggleContainer) {
                devToggleContainer.style.display = 'flex';
            }
            scriptState.container.classList.remove('hidden');
            console.log('[深大成绩查询] 开发者模式已启用，可以在界面中查看原始数据');
        }
    });

})();
