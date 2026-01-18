// ==UserScript==
// @name         深圳大学平时成绩&期末成绩查询
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  双向固定分区遍历，更快查询速度，优化UI显示，支持导出Excel
// @author       流年.
// @match        https://ehall.szu.edu.cn/jwapp/sys/cjcx/*
// @match        https://ehall-443.webvpn.szu.edu.cn/jwapp/sys/cjcx/*
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
            initialCourses: null
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
                    const pscjxs = parseFloat(course.PSCJXS) || 0;
                    const qmcjxs = parseFloat(course.QMCJXS) || 0;
                    
                    // 根据系数判断是否需要查询
                    // 如果平时成绩系数为0或null，则不需要查询平时成绩
                    // 如果期末成绩系数为0或null，则不需要查询期末成绩
                    const needPscj = pscjxs > 0;
                    const needQmcj = qmcjxs > 0;
                    
                    course.PSCJ = needPscj ? 'N/A' : '-';  // '-' 表示不需要查询
                    course.QMCJ = needQmcj ? 'N/A' : '-';
                    course.PSCJXS = course.PSCJXS || '0';
                    course.QMCJXS = course.QMCJXS || '0';
                    course._needPscj = needPscj;  // 内部标记
                    course._needQmcj = needQmcj;
                    
                    if (needPscj) needPscjCount++;
                    if (needQmcj) needQmcjCount++;
                    
                    courseMap.set(key, course);
                });

                console.log(`[深大成绩查询] 需要查询平时成绩: ${needPscjCount} 门, 期末成绩: ${needQmcjCount} 门`);

                let pscjFoundCount = 0;
                let qmcjFoundCount = 0;
                
                statusEl.textContent = '正在快速查询详细成绩...';

                // 3. 双向固定分区遍历（无需判断交叉，更高效）
                // 高分端: 100→51 (共50个分数)
                // 低分端: 0→50 (共51个分数)
                let highScore = 100;
                let lowScore = 0;
                const HIGH_END = 51;  // 高分端终点（含）
                const LOW_END = 50;   // 低分端终点（含）
                
                while (highScore >= HIGH_END || lowScore <= LOW_END) {
                    // 检查是否所有需要查询的成绩都已找到
                    const pscjDone = pscjFoundCount >= needPscjCount;
                    const qmcjDone = qmcjFoundCount >= needQmcjCount;
                    
                    if (pscjDone && qmcjDone) {
                        console.log('[深大成绩查询] 所有成绩已找到，提前结束');
                        break;
                    }

                    // 计算进度
                    const highProgress = 100 - highScore;  // 高分端已完成数 (0→49)
                    const lowProgress = lowScore;          // 低分端已完成数 (0→50)
                    const progress = Math.min(((highProgress + lowProgress) / 101) * 100, 100);
                    progressEl.style.width = `${progress}%`;
                    statusEl.textContent = `查询进度: ${Math.round(progress)}% (↓${highScore} ↑${lowScore}) [平时:${pscjFoundCount}/${needPscjCount} 期末:${qmcjFoundCount}/${needQmcjCount}]`;

                    // 并行查询：高分端和低分端同时进行
                    const queries = [];
                    
                    // 高分端查询 (100→51)
                    if (highScore >= HIGH_END) {
                        if (!pscjDone) {
                            queries.push(performQuery(highScore, 'PSCJ').then(rows => ({ type: 'PSCJ', score: highScore, rows })));
                        }
                        if (!qmcjDone) {
                            queries.push(performQuery(highScore, 'QMCJ').then(rows => ({ type: 'QMCJ', score: highScore, rows })));
                        }
                    }
                    
                    // 低分端查询 (0→50)
                    if (lowScore <= LOW_END) {
                        if (!pscjDone) {
                            queries.push(performQuery(lowScore, 'PSCJ').then(rows => ({ type: 'PSCJ', score: lowScore, rows })));
                        }
                        if (!qmcjDone) {
                            queries.push(performQuery(lowScore, 'QMCJ').then(rows => ({ type: 'QMCJ', score: lowScore, rows })));
                        }
                    }

                    // 等待所有并行查询完成
                    const results = await Promise.all(queries);

                    // 处理查询结果
                    results.forEach(result => {
                        result.rows.forEach(row => {
                            const key = row.KCM + row.XNXQDM_DISPLAY;
                            const course = courseMap.get(key);
                            if (course) {
                                if (result.type === 'PSCJ' && course.PSCJ === 'N/A' && course._needPscj) {
                                    course.PSCJ = result.score.toString();
                                    course.PSCJXS = row.PSCJXS || course.PSCJXS;
                                    pscjFoundCount++;
                                } else if (result.type === 'QMCJ' && course.QMCJ === 'N/A' && course._needQmcj) {
                                    course.QMCJ = result.score.toString();
                                    course.QMCJXS = row.QMCJXS || course.QMCJXS;
                                    qmcjFoundCount++;
                                }
                            }
                        });
                    });

                    // 更新数据和渲染
                    scriptState.courseData = Array.from(courseMap.values());
                    renderResults();

                    // 移动指针（独立移动，无需判断交叉）
                    if (highScore >= HIGH_END) highScore--;
                    if (lowScore <= LOW_END) lowScore++;

                    // 更短的延迟，加快查询速度（30ms）
                    await new Promise(resolve => setTimeout(resolve, 30));
                }

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

            // 准备表头（与前端展示的数据一致）
            const header = [
                '学期', '课程号', '课程名称', '课程类别', '开课学院', '课程学分',
                '平时成绩', '平时系数(%)', '期末成绩', '期末系数(%)',
                '总成绩', '等级', '等级制成绩'
            ];

            // 准备数据行
            const dataRows = scriptState.courseData.map(course => {
                const { finalScore, grade } = calculateFinalScoreAndGrade(course);
                return [
                    course.XNXQDM_DISPLAY || 'N/A',
                    course.KCH || 'N/A',
                    course.KCM || 'N/A',
                    course.KCLBDM_DISPLAY || 'N/A',
                    course.KKDWDM_DISPLAY || 'N/A',
                    course.XF || 'N/A',
                    course.PSCJ,
                    course.PSCJXS || 'N/A',
                    course.QMCJ,
                    course.QMCJXS || 'N/A',
                    finalScore,
                    grade,
                    course.XFJD || 'N/A'
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
                { wch: 12 }     // 等级制成绩
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
        const pscjxs = parseFloat(course.PSCJXS) || 0;
        const qmcjxs = parseFloat(course.QMCJXS) || 0;
        
        // 解析成绩，'-' 表示不需要该成绩，视为0分0权重
        const pscj = course.PSCJ === '-' ? 0 : parseFloat(course.PSCJ);
        const qmcj = course.QMCJ === '-' ? 0 : parseFloat(course.QMCJ);
        
        // 计算有效权重
        const effectivePscjxs = course.PSCJ === '-' ? 0 : pscjxs;
        const effectiveQmcjxs = course.QMCJ === '-' ? 0 : qmcjxs;

        let rawFinalScore;

        // 如果只有平时成绩（系数100%或期末系数为0）
        if (effectivePscjxs === 100 || (effectivePscjxs > 0 && effectiveQmcjxs === 0)) {
            if (!isNaN(pscj) && course.PSCJ !== 'N/A') {
                rawFinalScore = pscj;
            } else {
                // 还没查到平时成绩
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        // 如果只有期末成绩（期末系数100%或平时系数为0）
        else if (effectiveQmcjxs === 100 || (effectiveQmcjxs > 0 && effectivePscjxs === 0)) {
            if (!isNaN(qmcj) && course.QMCJ !== 'N/A') {
                rawFinalScore = qmcj;
            } else {
                // 还没查到期末成绩
                if (course.ZCJ != null) {
                    return { finalScore: course.ZCJ, grade: course.DJCJMC || 'N/A' };
                }
                return { finalScore: 'N/A', grade: 'N/A' };
            }
        }
        // 正常情况：平时+期末
        else if (!isNaN(pscj) && !isNaN(qmcj) && course.PSCJ !== 'N/A' && course.QMCJ !== 'N/A') {
            rawFinalScore = (pscj * pscjxs / 100) + (qmcj * qmcjxs / 100);
        } else {
            // 成绩不完整，使用服务器返回的总成绩
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
                        <span>平时: <b style="color: #4CAF50;">${course.PSCJ}</b> (${course.PSCJXS}%)</span>
                        <span>期末: <b style="color: #FF5722;">${course.QMCJ}</b> (${course.QMCJXS}%)</span>
                    </div>
                    
                    <div class="course-detail full-width score-row" style="margin-top: 4px; padding-top: 4px; border-top: 1px solid #eee;">
                        <span>总评: <span class="final-score">${finalScore}</span> <span class="final-score">(${grade})</span></span>
                    </div>
                `;
                resultsEl.appendChild(item);
            });
        });
    }

    // 更新开发者模式数据显示
    function updateDevDataDisplay() {
        if (!scriptState.container) return;
        
        const initialDataEl = scriptState.container.querySelector('#dev-initial-data');
        
        if (initialDataEl && scriptState.rawData.initialCourses !== null) {
            initialDataEl.textContent = JSON.stringify(scriptState.rawData.initialCourses, null, 2);
        }
    }

    toggleBtn.addEventListener('click', () => scriptState.container.classList.toggle('hidden'));

    // 获取初始课程列表
    function fetchInitialCourseList() {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: "GET",
                url: `${location.origin}/jwapp/sys/cjcx/modules/cjcx/xscjcx.do`,
                headers: { "Cookie": document.cookie },
                onload: res => {
                    try {
                        const data = JSON.parse(res.responseText);
                        scriptState.rawData.initialCourses = data;
                        if (scriptState.devMode) {
                            updateDevDataDisplay();
                        }
                        resolve(data?.datas?.xscjcx?.rows || []);
                    } catch (e) { reject(new Error("解析初始课程列表失败")); }
                },
                onerror: () => reject(new Error("获取初始课程列表网络请求失败"))
            });
        });
    }

    // 执行成绩查询
    function performQuery(score, scoreType) {
        return new Promise(resolve => {
            const payload = `querySetting=[{"name":"${scoreType}","value":"${score}","linkOpt":"and","builder":"equal"}]&pageSize=100&pageNumber=1`;
            GM_xmlhttpRequest({
                method: "POST",
                url: `${location.origin}/jwapp/sys/cjcx/modules/cjcx/xscjcx.do`,
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
                    "Cookie": document.cookie
                },
                data: payload,
                onload: res => {
                    try {
                        const data = JSON.parse(res.responseText);
                        resolve(data?.datas?.xscjcx?.rows || []);
                    } catch (e) {
                        console.error(`解析${scoreType}=${score}的响应失败:`, e);
                        resolve([]);
                    }
                },
                onerror: () => {
                    console.error(`查询${scoreType}=${score}时网络请求失败`);
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
