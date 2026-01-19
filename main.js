// ==UserScript==
// @name         深圳大学平时成绩&期末成绩查询
// @namespace    http://tampermonkey.net/
// @version      4.6
// @description  10线程并行分段查询，自动推算系数（支持0:100），优化UI显示，支持导出Excel
// @author       流年.
// @match        https://ehall.szu.edu.cn/*
// @match        https://ehall-443.webvpn.szu.edu.cn/*
// @connect      ehall.szu.edu.cn
// @connect      ehall-443.webvpn.szu.edu.cn
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        GM_registerMenuCommand
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
        rawData: {
            initialCourses: null,
            queryResults: []  // 存储轮询结果
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
            overflow-y: auto;
            margin: 0 -12px;
            padding: 4px 12px;
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

        startBtn.addEventListener('click', async () => {
            if (scriptState.isRunning) return;

            getStudentInfoFromPage();

            scriptState.isRunning = true;
            startBtn.disabled = true;
            exportBtn.disabled = true;
            resultsEl.innerHTML = '';
            progressEl.style.width = '0%';
            // 显示进度条区域
            const progressContainer = container.querySelector('.progress-container');
            progressContainer.classList.remove('completed');
            progressContainer.classList.add('active');
            statusEl.textContent = '正在获取课程列表...';

            try {
                // 1. 获取初始课程列表
                const initialCourses = await fetchInitialCourseList();
                if (!initialCourses || initialCourses.length === 0) {
                    statusEl.textContent = '未找到任何课程记录，请确认当前学期有成绩。';
                    return;
                }

                // 2. 初始化课程Map，并根据系数判断需要查询哪些成绩
                const courseMap = new Map();
                let needPscjCount = 0;  // 需要查询平时成绩的课程数
                let needQmcjCount = 0;  // 需要查询期末成绩的课程数
                
                initialCourses.forEach(course => {
                    const key = course.KCM + course.XNXQDM_DISPLAY;
                    
                    // 接口已不返回系数，需要查询两个成绩后自动计算
                    // 初始化时设置为需要查询
                    course.PSCJ = 'N/A';
                    course.QMCJ = 'N/A';
                    course.PSCJXS = '?';  // '?' 表示待计算
                    course.QMCJXS = '?';
                    
                    // 内部使用的数值系数（初始为null，待推算）
                    course._pscjxsNum = null;
                    course._qmcjxsNum = null;
                    course._needPscj = true;
                    course._needQmcj = true;
                    course._coefficientsInferred = false;
                    
                    // 保存原始总成绩用于后续推算系数
                    course._originalZCJ = course.ZCJ;
                    
                    needPscjCount++;
                    needQmcjCount++;
                    
                    courseMap.set(key, course);
                });

                console.log(`[深大成绩查询] 需要查询平时成绩: ${needPscjCount} 门, 期末成绩: ${needQmcjCount} 门`);

                let pscjFoundCount = 0;
                let qmcjFoundCount = 0;
                
                statusEl.textContent = '正在查询详细成绩...';

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
                    const progress = Math.min((sharedState.queriedScores.size / totalScores) * 100, 100);
                    progressEl.style.width = `${progress}%`;
                    statusEl.textContent = `并行查询中... [平时:${sharedState.pscjFoundCount}/${needPscjCount} 期末:${sharedState.qmcjFoundCount}/${needQmcjCount}] (已查${sharedState.queriedScores.size}个分数)`;
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
                    if (course._coefficientsInferred) {
                        return; // 已经推算过
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

                progressEl.style.width = '100%';
                statusEl.textContent = `查询完成！共 ${courseMap.size} 门课程`;
                // 查询完成后隐藏进度条区域
                container.querySelector('.progress-container').classList.add('completed');
                exportBtn.disabled = false;

            } catch (err) {
                console.error("查询过程中发生错误:", err);
                statusEl.textContent = `查询异常: ${err.message}`;
            } finally {
                scriptState.isRunning = false;
                startBtn.disabled = false;
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
                if (course._coefficientsInferred) {
                    coefficientSource = '推算';
                } else if (course.PSCJXS && !course.PSCJXS.endsWith('*') && course.PSCJXS !== '?') {
                    coefficientSource = '接口返回';
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
                <div style="margin-bottom:8px;">
                    <div style="font-size:0.8rem;color:#666;margin-bottom:4px;font-weight:500;">${title}</div>
                    <svg width="${chartWidth}" height="${chartHeight}" style="background:#fff;border-radius:6px;border:1px solid #e0e0e0;">
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

        let html = '<div style="margin-top:8px;">';
        
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

    function renderResults() {
        const resultsEl = scriptState.container.querySelector('#score-results');
        resultsEl.innerHTML = '';
        
        const courses = scriptState.courseData;
        if (courses.length === 0) {
            resultsEl.innerHTML = '<div style="text-align:center;padding:20px;color:#999;">暂无数据</div>';
            return;
        }

        // 1. 计算总 GPA
        const totalGPA = calculateGPA(courses);

        // 2. 计算学年 GPA
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

        // 4. 渲染概览区域
        const summaryDiv = document.createElement('div');
        summaryDiv.style.cssText = 'background:#e3f2fd;padding:12px;border-radius:8px;margin-bottom:16px;border:1px solid #bbdefb;';
        let summaryHTML = `<div style="font-size:1.1rem;font-weight:bold;color:#1565c0;margin-bottom:8px;">总 GPA: ${totalGPA}</div>`;
        summaryHTML += `<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px;">`;
        yearGPAs.forEach(item => {
            summaryHTML += `<span style="background:#fff;padding:4px 8px;border-radius:4px;font-size:0.85rem;color:#555;border:1px solid #e0e0e0;">${item.year}学年: <b>${item.gpa}</b></span>`;
        });
        summaryHTML += `</div>`;
        
        // 5. 添加 GPA 趋势图
        summaryHTML += renderGPAChart(semesterGPAData, yearGPAs);
        
        summaryDiv.innerHTML = summaryHTML;
        resultsEl.appendChild(summaryDiv);

        // 4. 按学期分组并渲染
        const sortedCourses = [...courses].sort((a, b) => {
            if (a.XNXQDM !== b.XNXQDM) {
                return (b.XNXQDM || '').localeCompare(a.XNXQDM || '');
            }
            return a.KCM.localeCompare(b.KCM);
        });

        const semesterGroups = new Map();
        sortedCourses.forEach(course => {
            const key = course.XNXQDM_DISPLAY || '未知学期';
            if (!semesterGroups.has(key)) {
                semesterGroups.set(key, []);
            }
            semesterGroups.get(key).push(course);
        });

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
        if (xs === '?') return '?';
        if (xs.endsWith('*')) {
            // 推断值，显示带提示
            return xs.replace('*', '') + '% (推断)';
        }
        return xs + '%';
    }

    // 更新开发者模式数据显示
    function updateDevDataDisplay() {
        if (!scriptState.container) return;
        
        const initialDataEl = scriptState.container.querySelector('#dev-initial-data');
        
        if (initialDataEl && scriptState.rawData.initialCourses !== null) {
            initialDataEl.textContent = JSON.stringify(scriptState.rawData.initialCourses, null, 2);
        }
        
        updateDevQueryDisplay();
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
