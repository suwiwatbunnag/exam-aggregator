document.addEventListener('DOMContentLoaded', () => {
    // Initialize Lucide icons
    lucide.createIcons();

    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileListEl = document.getElementById('file-list');
    const configSection = document.getElementById('config-section');
    const resultsSection = document.getElementById('results-section');
    const chartSection = document.getElementById('chart-section');
    const moduleConfigGrid = document.getElementById('module-config-grid');
    const resultsTable = document.getElementById('results-table');
    const tableHeader = document.getElementById('table-header');
    const tableBody = document.getElementById('table-body');
    
    const feedbackSection = document.getElementById('feedback-section');
    const feedbackHeader = document.getElementById('feedback-header');
    const feedbackBody = document.getElementById('feedback-body');
    
    const exportScoresBtn = document.getElementById('export-scores-btn');
    const exportFeedbackBtn = document.getElementById('export-feedback-btn');
    const exportGradingBtn = document.getElementById('export-grading-btn');
    const exportStatsBtn = document.getElementById('export-stats-btn');
    const downloadChartBtn = document.getElementById('download-chart-btn');
    
    const gradingSection = document.getElementById('grading-section');
    const gradingHeader = document.getElementById('grading-header');
    const gradingBody = document.getElementById('grading-body');

    let allStudents = {}; // Keyed by username
    let modulesFound = {}; // Keyed by module name
    let uploadedFiles = [];
    let currentSort = { column: '#', direction: 'asc' }; // Sorting state
    let incrementalRate = 4; // Default 4%
    let gradingScenarios = {
        1: 4.0,
        2: 3.0,
        3: 3.5
    };
    let activeScenario = 1;

    // Drag and Drop handlers
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files);
    });

    fileInput.addEventListener('change', (e) => handleFiles(e.target.files));

    async function handleFiles(files) {
        if (files.length === 0) return;
        
        uploadedFiles = Array.from(files);
        fileListEl.innerHTML = '';
        allStudents = {};
        modulesFound = {};

        for (const file of uploadedFiles) {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `<span><i data-lucide="file-text"></i> ${file.name}</span>`;
            fileListEl.appendChild(item);
            await processFile(file);
        }
        lucide.createIcons();
        renderConfig();
        updateAggregates();
    }

    async function processFile(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                jsonData.forEach(row => {
                    const username = row['username'] || row['Username'];
                    if (!username) return;

                    if (!allStudents[username]) {
                        allStudents[username] = {
                            '#': row['#'],
                            'username': username,
                            'firstname': row['firstname'] || row['Firstname'],
                            'surname': row['surname'] || row['Surname'],
                            'scores': {}
                        };
                    }

                    // Look for module column (xxxx_classyyyy)
                    Object.keys(row).forEach(key => {
                        if (key.includes('_class')) {
                            const moduleName = key;
                            const score = parseFloat(row[key]);
                            allStudents[username].scores[moduleName] = isNaN(score) ? 0 : score;

                            if (!modulesFound[moduleName]) {
                                modulesFound[moduleName] = {
                                    name: moduleName,
                                    fullScore: 40, // Default
                                    mpl: 0.4 // Default
                                };
                            }
                        }
                    });
                });
                resolve();
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function renderConfig() {
        const modules = Object.keys(modulesFound);
        if (modules.length === 0) return;

        configSection.classList.remove('hidden');
        moduleConfigGrid.innerHTML = '';

        modules.forEach((modName, index) => {
            const mod = modulesFound[modName];
            const colorVar = `--mod-${(index % 4) + 1}`;
            
            const card = document.createElement('div');
            card.className = 'module-card';
            card.style.borderLeftColor = `var(${colorVar})`;
            card.innerHTML = `
                <h3>${modName}</h3>
                <div class="input-group">
                    <label>Full Score / Total Qs</label>
                    <input type="number" value="${mod.fullScore}" data-mod="${modName}" data-type="fullScore">
                </div>
                <div class="input-group">
                    <label>MPL (e.g. 0.6 for 60%)</label>
                    <input type="number" step="0.01" value="${mod.mpl}" data-mod="${modName}" data-type="mpl">
                </div>
                <div class="pass-info" id="pass-info-${modName}">
                    Calculated Pass: <span class="raw">0</span> | Ceil: <span class="ceil">0</span>
                </div>
            `;
            moduleConfigGrid.appendChild(card);
        });

        // Add listeners to inputs
        moduleConfigGrid.querySelectorAll('input').forEach(input => {
            input.addEventListener('input', (e) => {
                const modName = e.target.dataset.mod;
                const type = e.target.dataset.type;
                modulesFound[modName][type] = parseFloat(e.target.value) || 0;
                updateAggregates();
            });
        });

        document.getElementById('stat-count').innerText = `${modules.length} / 4`;
    }

    function updateAggregates() {
        let totalFullScore = 0;
        let weightedMPLSum = 0;

        Object.keys(modulesFound).forEach(modName => {
            const mod = modulesFound[modName];
            const rawPass = mod.fullScore * mod.mpl;
            const ceilPass = Math.ceil(rawPass);
            
            mod.rawPassMark = rawPass;
            mod.ceilPassMark = ceilPass;

            const infoEl = document.getElementById(`pass-info-${modName}`);
            if (infoEl) {
                infoEl.querySelector('.raw').innerText = rawPass.toFixed(2);
                infoEl.querySelector('.ceil').innerText = ceilPass;
            }

            totalFullScore += mod.fullScore;
            weightedMPLSum += (mod.mpl * mod.fullScore);
        });

        const avgMPL = totalFullScore > 0 ? (weightedMPLSum / totalFullScore) : 0;
        const totalPassMark = avgMPL * totalFullScore;
        const totalPassMarkCeil = Math.ceil(totalPassMark);

        document.getElementById('stat-avg-mpl').innerText = (avgMPL * 100).toFixed(2) + '%';
        document.getElementById('stat-total-pass').innerText = `${totalPassMark.toFixed(2)} (Ceil: ${totalPassMarkCeil})`;

        // Store for global use
        window.avgMPL = avgMPL;
        window.totalPassMarkCeil = totalPassMarkCeil;
        window.totalFullScore = totalFullScore;

        renderResultsTable();
        renderFeedbackTable();
        renderGradingSummaries();
        renderGradingTable();
        renderBoxPlot();
        renderStatsTable();
    }

    const incrementalRateInput = document.getElementById('incremental-rate');
    if (incrementalRateInput) {
        incrementalRateInput.addEventListener('input', (e) => {
            incrementalRate = parseFloat(e.target.value) || 0;
            updateAggregates();
        });
    }

    // Grading Scenario Handlers
    document.querySelectorAll('.scenario-step').forEach(input => {
        input.addEventListener('input', (e) => {
            const id = e.target.dataset.id;
            gradingScenarios[id] = parseFloat(e.target.value) || 0;
            renderGradingSummaries();
            if (activeScenario == id) renderGradingTable();
        });
    });

    document.querySelectorAll('.apply-scenario-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            activeScenario = btn.dataset.id;
            document.getElementById('active-scenario-label').textContent = activeScenario;
            
            // Highlight active button
            document.querySelectorAll('.apply-scenario-btn').forEach(b => {
                b.classList.add('secondary-style');
            });
            btn.classList.remove('secondary-style');
            
            renderGradingTable();
        });
    });

    function calculateGrade(percent, mplPercent, step) {
        if (percent < mplPercent) return 'D/F';
        if (percent < mplPercent + step) return 'D+';
        if (percent < mplPercent + 2 * step) return 'C';
        if (percent < mplPercent + 3 * step) return 'C+';
        if (percent < mplPercent + 4 * step) return 'B';
        if (percent < mplPercent + 5 * step) return 'B+';
        return 'A';
    }

    function renderGradingSummaries() {
        const grades = ['D/F', 'D+', 'C', 'C+', 'B', 'B+', 'A'];
        const totalFullScore = window.totalFullScore || 100;
        const mplPercent = (window.avgMPL || 0) * 100;

        [1, 2, 3].forEach(id => {
            const step = gradingScenarios[id];
            const counts = { 'D/F': 0, 'D+': 0, 'C': 0, 'C+': 0, 'B': 0, 'B+': 0, 'A': 0 };
            
            // Count students
            Object.values(allStudents).forEach(student => {
                let total = 0;
                Object.keys(modulesFound).forEach(m => total += (student.scores[m] || 0));
                const p = (total / totalFullScore) * 100;
                counts[calculateGrade(p, mplPercent, step)]++;
            });

            const body = document.getElementById(`summary-body-${id}`);
            if (!body) return;

            let html = '';
            
            // Row 1: Score (%)
            html += `<tr><td>Min % Score</td>`;
            grades.forEach((g, idx) => {
                if (g === 'D/F') html += `<td>0.00%</td>`;
                else if (g === 'D+') html += `<td>${mplPercent.toFixed(2)}%</td>`;
                else html += `<td>${(mplPercent + (idx - 1) * step).toFixed(2)}%</td>`;
            });
            html += `</tr>`;

            // Row 2: Score (Points)
            html += `<tr><td>Min Points</td>`;
            grades.forEach((g, idx) => {
                if (g === 'D/F') html += `<td>0.00</td>`;
                else if (g === 'D+') html += `<td>${(totalFullScore * mplPercent / 100).toFixed(2)}</td>`;
                else html += `<td>${(totalFullScore * (mplPercent + (idx - 1) * step) / 100).toFixed(2)}</td>`;
            });
            html += `</tr>`;

            // Row 3: Count
            html += `<tr><td>No. of Students</td>`;
            grades.forEach(g => html += `<td>${counts[g]}</td>`);
            html += `</tr>`;

            // Row 4: % of Class
            const totalStudents = Object.keys(allStudents).length || 1;
            html += `<tr><td>% of Class</td>`;
            grades.forEach(g => html += `<td>${((counts[g] / totalStudents) * 100).toFixed(1)}%</td>`);
            html += `</tr>`;

            body.innerHTML = html;
        });
    }

    function renderGradingTable() {
        if (!gradingSection || !gradingBody) return;
        const modules = Object.keys(modulesFound);
        if (modules.length === 0) return;

        gradingSection.classList.remove('hidden');
        const step = gradingScenarios[activeScenario];
        const mplPercent = (window.avgMPL || 0) * 100;
        
        // Header
        gradingHeader.innerHTML = `<th>#</th><th>Username</th><th>Firstname</th><th>Surname</th><th>Total Score</th><th>Percentage (%)</th><th>Grade</th>`;

        // Body
        gradingBody.innerHTML = '';
        Object.values(allStudents).forEach(student => {
            const tr = document.createElement('tr');
            let totalScore = 0;
            modules.forEach(modName => totalScore += (student.scores[modName] || 0));
            
            const percent = (totalScore / window.totalFullScore) * 100;
            const grade = calculateGrade(percent, mplPercent, step);
            
            tr.innerHTML = `
                <td>${student['#']}</td>
                <td>${student['username']}</td>
                <td>${student['firstname']}</td>
                <td>${student['surname']}</td>
                <td>${totalScore.toFixed(2)}</td>
                <td>${percent.toFixed(2)}%</td>
                <td style="font-weight: bold; color: var(--primary);">${grade}</td>
            `;
            gradingBody.appendChild(tr);
        });
    }

    function getFeedbackCategory(score, fullScore, mpl) {
        if (!fullScore || fullScore === 0) return { label: 'N/A', class: '' };
        
        const scorePercent = (score / fullScore) * 100;
        const mplPercent = mpl * 100;
        const ir = incrementalRate;

        if (scorePercent >= mplPercent + 20) return { label: 'Very good', class: 'fb-very-good' };
        if (scorePercent >= mplPercent + 12) return { label: 'Good', class: 'fb-good' };
        if (scorePercent >= mplPercent + 4) return { label: 'Borderline', class: 'fb-borderline' };
        return { label: 'Needs improvement', class: 'fb-needs-improvement' };
    }

    function calculateStats(scores) {
        if (scores.length === 0) return { min: 0, max: 0, mean: 0, sd: 0, median: 0, q1: 0, q3: 0, iqr: 0 };
        const sorted = [...scores].sort((a, b) => a - b);
        const min = sorted[0];
        const max = sorted[sorted.length - 1];
        const sum = sorted.reduce((a, b) => a + b, 0);
        const mean = sum / sorted.length;
        const sd = Math.sqrt(sorted.map(x => Math.pow(x - mean, 2)).reduce((a, b) => a + b, 0) / sorted.length);
        const median = sorted[Math.floor(sorted.length / 2)];
        const q1 = sorted[Math.floor(sorted.length / 4)];
        const q3 = sorted[Math.floor(sorted.length * 0.75)];
        const iqr = q3 - q1;
        return { min, max, mean, sd, median, q1, q3, iqr };
    }

    function renderStatsTable() {
        const modules = Object.keys(modulesFound);
        const statsHeader = document.getElementById('stats-header');
        const statsBody = document.getElementById('stats-body');

        let headerHtml = '<th>Metric</th>';
        modules.forEach(m => headerHtml += `<th>${m}</th>`);
        headerHtml += `<th>TOTAL</th>`;
        statsHeader.innerHTML = headerHtml;

        const metrics = [
            { label: 'Max', key: 'max' },
            { label: 'Min', key: 'min' },
            { label: 'Mean (SD)', format: s => `${s.mean.toFixed(2)} (${s.sd.toFixed(2)})` },
            { label: 'Median (IQR)', format: s => `${s.median.toFixed(2)} (${s.iqr.toFixed(2)})` }
        ];

        statsBody.innerHTML = '';
        metrics.forEach(metric => {
            const tr = document.createElement('tr');
            let rowHtml = `<td><strong>${metric.label}</strong></td>`;
            
            const allModStats = [];
            modules.forEach(modName => {
                const scores = Object.values(allStudents).map(s => s.scores[modName] || 0);
                const s = calculateStats(scores);
                allModStats.push(s);
                rowHtml += `<td>${metric.format ? metric.format(s) : s[metric.key].toFixed(2)}</td>`;
            });

            // Total Stats
            const totalScores = Object.values(allStudents).map(student => {
                return Object.keys(modulesFound).reduce((sum, mod) => sum + (student.scores[mod] || 0), 0);
            });
            const ts = calculateStats(totalScores);
            rowHtml += `<td>${metric.format ? metric.format(ts) : ts[metric.key].toFixed(2)}</td>`;

            tr.innerHTML = rowHtml;
            statsBody.appendChild(tr);
        });
    }

    function renderResultsTable() {
        const modules = Object.keys(modulesFound);
        if (modules.length === 0) return;

        resultsSection.classList.remove('hidden');
        
        // Header
        let headerHtml = `<th class="sortable" data-col="#"># <i data-lucide="arrow-up-down" class="sort-icon"></i></th><th>Username</th><th>Firstname</th><th>Surname</th>`;
        modules.forEach(mod => {
            headerHtml += `<th class="sortable" data-col="${mod}">${mod} <i data-lucide="arrow-up-down" class="sort-icon"></i></th>`;
        });
        headerHtml += `<th class="sortable" data-col="sum">SUM (${window.totalFullScore}) <i data-lucide="arrow-up-down" class="sort-icon"></i></th>`;
        tableHeader.innerHTML = headerHtml;

        // Add sorting listeners
        tableHeader.querySelectorAll('th.sortable').forEach(th => {
            th.addEventListener('click', () => {
                const col = th.dataset.col;
                if (currentSort.column === col) {
                    currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
                } else {
                    currentSort.column = col;
                    currentSort.direction = 'asc';
                }
                renderResultsTable();
            });
        });
        lucide.createIcons();

        // Sort Data
        const sortedStudents = Object.values(allStudents).sort((a, b) => {
            let valA, valB;
            if (currentSort.column === '#') {
                valA = parseInt(a['#']) || 0;
                valB = parseInt(b['#']) || 0;
            } else if (currentSort.column === 'sum') {
                valA = Object.values(a.scores).reduce((sum, s) => sum + s, 0);
                valB = Object.values(b.scores).reduce((sum, s) => sum + s, 0);
            } else {
                valA = a.scores[currentSort.column] || 0;
                valB = b.scores[currentSort.column] || 0;
            }
            
            if (currentSort.direction === 'asc') return valA - valB;
            return valB - valA;
        });

        // Body
        tableBody.innerHTML = '';
        sortedStudents.forEach(student => {
            const tr = document.createElement('tr');
            let rowSum = 0;
            
            let rowHtml = `
                <td>${student['#']}</td>
                <td>${student['username']}</td>
                <td>${student['firstname']}</td>
                <td>${student['surname']}</td>
            `;

            modules.forEach((modName, idx) => {
                const score = student.scores[modName] || 0;
                rowSum += score;
                const mod = modulesFound[modName];
                const isFail = score < mod.ceilPassMark;
                const failClass = isFail ? `fail-mod-${(idx % 4) + 1}` : '';
                rowHtml += `<td class="${failClass}">${score}</td>`;
            });

            const isTotalFail = rowSum < window.totalPassMarkCeil;
            rowHtml += `<td class="${isTotalFail ? 'fail-sum' : ''}">${rowSum}</td>`;
            
            tr.innerHTML = rowHtml;
            tableBody.appendChild(tr);
        });
    }

    function renderFeedbackTable() {
        const modules = Object.keys(modulesFound);
        if (modules.length === 0) return;

        feedbackSection.classList.remove('hidden');
        
        // Header
        let headerHtml = `<th>#</th><th>Username</th><th>Firstname</th><th>Surname</th>`;
        modules.forEach(mod => headerHtml += `<th>${mod} Feedback</th>`);
        feedbackHeader.innerHTML = headerHtml;

        // Body
        feedbackBody.innerHTML = '';
        Object.values(allStudents).forEach(student => {
            const tr = document.createElement('tr');
            let rowSum = 0;
            
            let rowHtml = `
                <td>${student['#']}</td>
                <td>${student['username']}</td>
                <td>${student['firstname']}</td>
                <td>${student['surname']}</td>
            `;

            modules.forEach((modName) => {
                const score = student.scores[modName] || 0;
                rowSum += score;
                const mod = modulesFound[modName];
                const fb = getFeedbackCategory(score, mod.fullScore, mod.mpl);
                rowHtml += `<td><span class="feedback-badge ${fb.class}">${fb.label}</span></td>`;
            });

            tr.innerHTML = rowHtml;
            feedbackBody.appendChild(tr);
        });
    }

    function renderBoxPlot() {
        chartSection.classList.remove('hidden');
        const modules = Object.keys(modulesFound);
        const plotData = [];
        const colors = ["#D62728", "#1F77B4", "#2CA02C", "#9467BD", "#FF7F0E"];

        // 1. Module Traces
        modules.forEach((modName, idx) => {
            const mod = modulesFound[modName];
            const scores = Object.values(allStudents).map(s => ((s.scores[modName] || 0) / mod.fullScore) * 100);
            
            // Box + Swarm
            plotData.push({
                y: scores,
                type: 'box',
                name: modName,
                boxpoints: 'all',
                jitter: 0.5,
                pointpos: 0,
                marker: { color: colors[idx % colors.length], size: 4, opacity: 0.7 },
                fillcolor: colors[idx % colors.length],
                line: { color: colors[idx % colors.length] },
                opacity: 0.6,
                hoverinfo: 'y+name'
            });

            // MPL X Mark
            plotData.push({
                x: [modName],
                y: [mod.mpl * 100],
                mode: 'markers',
                marker: { symbol: 'x', size: 12, color: 'red', line: { width: 2 } },
                name: 'MPL',
                showlegend: idx === 0,
                hoverinfo: 'none'
            });
        });

        // 2. Total Trace
        const totalScores = Object.values(allStudents).map(student => {
            let sum = 0;
            modules.forEach(modName => sum += (student.scores[modName] || 0));
            return (sum / window.totalFullScore) * 100;
        });

        plotData.push({
            y: totalScores,
            type: 'box',
            name: 'Total',
            boxpoints: 'all',
            jitter: 0.5,
            pointpos: 0,
            marker: { color: '#FF7F0E', size: 4, opacity: 0.7 },
            fillcolor: '#FF7F0E',
            line: { color: '#FF7F0E' },
            opacity: 0.6,
            hoverinfo: 'y+name'
        });

        // Total MPL
        plotData.push({
            x: ['Total'],
            y: [window.avgMPL * 100],
            mode: 'markers',
            marker: { symbol: 'x', size: 12, color: 'red', line: { width: 2 } },
            name: 'Avg MPL',
            showlegend: false,
            hoverinfo: 'none'
        });

        const layout = {
            title: { text: 'Score Distribution by Topic with MPL', font: { color: '#000000', size: 20 } },
            paper_bgcolor: '#ffffff',
            plot_bgcolor: '#ffffff',
            yaxis: {
                title: 'Score (%)',
                range: [0, 110],
                gridcolor: 'rgba(0,0,0,0.1)',
                tickfont: { color: '#333333' },
                titlefont: { color: '#000000' }
            },
            xaxis: {
                tickfont: { color: '#333333' },
                gridcolor: 'rgba(0,0,0,0.05)'
            },
            showlegend: true,
            legend: { font: { color: '#333333' } },
            margin: { l: 50, r: 50, b: 50, t: 80 }
        };

        const config = { responsive: true, displayModeBar: false };

        // We use a div instead of canvas for Plotly
        const chartDiv = document.getElementById('performance-chart-container');
        if (!chartDiv) {
            const container = document.querySelector('.chart-container');
            container.innerHTML = '<div id="performance-chart-container" style="height: 500px;"></div>';
        }
        
        Plotly.newPlot('performance-chart-container', plotData, layout, config);
    }

    if (downloadChartBtn) {
        downloadChartBtn.addEventListener('click', () => {
            Plotly.downloadImage('performance-chart-container', {
                format: 'png',
                width: 1200,
                height: 800,
                filename: 'Score_Distribution_Chart'
            });
        });
    }

    if (exportStatsBtn) {
        exportStatsBtn.addEventListener('click', () => {
            const modules = Object.keys(modulesFound);
            const statsData = [];
            const metrics = ['Max', 'Min', 'Mean', 'SD', 'Median', 'IQR'];
            metrics.forEach(metric => {
                const row = { 'Metric': metric };
                modules.forEach(modName => {
                    const scores = Object.values(allStudents).map(s => s.scores[modName] || 0);
                    const s = calculateStats(scores);
                    row[modName] = metric === 'Mean' ? s.mean.toFixed(2) : 
                                  metric === 'SD' ? s.sd.toFixed(2) : 
                                  s[metric.toLowerCase()].toFixed(2);
                });
                const totalScores = Object.values(allStudents).map(student => {
                    return Object.keys(modulesFound).reduce((sum, m) => sum + (student.scores[m] || 0), 0);
                });
                const ts = calculateStats(totalScores);
                row['TOTAL'] = metric === 'Mean' ? ts.mean.toFixed(2) : 
                               metric === 'SD' ? ts.sd.toFixed(2) : 
                               ts[metric.toLowerCase()].toFixed(2);
                statsData.push(row);
            });
            const ws = XLSX.utils.json_to_sheet(statsData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Statistics");
            
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
            const uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
            const a = document.createElement("a");
            document.body.appendChild(a);
            a.style = "display: none";
            a.href = uri;
            a.download = "Exam_Statistics_Summary.xlsx";
            a.click();
            setTimeout(() => document.body.removeChild(a), 100);
        });
    }

    exportScoresBtn.addEventListener('click', () => {
        const modules = Object.keys(modulesFound);
        const data = [];
        
        Object.values(allStudents).forEach(student => {
            const row = {
                '#': student['#'],
                'Username': student['username'],
                'Firstname': student['firstname'],
                'Surname': student['surname']
            };
            let sum = 0;
            modules.forEach(modName => {
                const score = student.scores[modName] || 0;
                row[modName] = score;
                sum += score;
            });
            row[`SUM (${window.totalFullScore})`] = sum;
            data.push(row);
        });

        const ws = XLSX.utils.json_to_sheet(data);

        // --- Statistics Sheet ---
        const statsData = [];
        const metrics = ['Max', 'Min', 'Mean', 'SD', 'Median', 'IQR'];
        metrics.forEach(metric => {
            const row = { 'Metric': metric };
            modules.forEach(modName => {
                const scores = Object.values(allStudents).map(s => s.scores[modName] || 0);
                const s = calculateStats(scores);
                row[modName] = metric === 'Mean' ? s.mean.toFixed(2) : 
                              metric === 'SD' ? s.sd.toFixed(2) : 
                              s[metric.toLowerCase()].toFixed(2);
            });
            const totalScores = Object.values(allStudents).map(student => {
                return Object.keys(modulesFound).reduce((sum, m) => sum + (student.scores[m] || 0), 0);
            });
            const ts = calculateStats(totalScores);
            row['TOTAL'] = metric === 'Mean' ? ts.mean.toFixed(2) : 
                           metric === 'SD' ? ts.sd.toFixed(2) : 
                           ts[metric.toLowerCase()].toFixed(2);
            statsData.push(row);
        });
        const wsStats = XLSX.utils.json_to_sheet(statsData);

        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            const studentIdx = R - 1;
            const student = Object.values(allStudents)[studentIdx];
            let rowSum = 0;
            modules.forEach((modName, Cidx) => {
                const colIdx = 4 + Cidx;
                const score = student.scores[modName] || 0;
                rowSum += score;
                const mod = modulesFound[modName];
                if (score < mod.ceilPassMark) {
                    const colors = ['3b82f6', '10b981', 'f59e0b', 'ef4444'];
                    const cellRef = XLSX.utils.encode_cell({r: R, c: colIdx});
                    if (!ws[cellRef]) ws[cellRef] = { v: score };
                    ws[cellRef].s = { fill: { fgColor: { rgb: colors[Cidx % 4] } }, font: { color: { rgb: "FFFFFF" }, bold: true } };
                }
            });
            const sumColIdx = 4 + modules.length;
            if (rowSum < window.totalPassMarkCeil) {
                const cellRef = XLSX.utils.encode_cell({r: R, c: sumColIdx});
                if (!ws[cellRef]) ws[cellRef] = { v: rowSum };
                ws[cellRef].s = { fill: { fgColor: { rgb: "EC4899" } }, font: { color: { rgb: "FFFFFF" }, bold: true } };
            }
        }
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Scores");
        XLSX.utils.book_append_sheet(wb, wsStats, "Statistics Summary");
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
        const uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
        const a = document.createElement("a");
        document.body.appendChild(a);
        a.style = "display: none";
        a.href = uri;
        a.download = "Exam_Scores_and_Stats.xlsx";
        a.click();
        setTimeout(() => document.body.removeChild(a), 100);
    });

    exportFeedbackBtn.addEventListener('click', () => {
        const modules = Object.keys(modulesFound);
        const data = [];
        Object.values(allStudents).forEach(student => {
            const row = {
                '#': student['#'],
                'Username': student['username'],
                'Firstname': student['firstname'],
                'Surname': student['surname']
            };
            let sum = 0;
            modules.forEach(modName => {
                const score = student.scores[modName] || 0;
                const mod = modulesFound[modName];
                row[modName + ' Feedback'] = getFeedbackCategory(score, mod.fullScore, mod.mpl).label;
            });
            data.push(row);
        });
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Feedback");
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
        const uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
        const a = document.createElement("a");
        document.body.appendChild(a);
        a.style = "display: none";
        a.href = uri;
        a.download = "Exam_Feedback.xlsx";
        a.click();
        setTimeout(() => document.body.removeChild(a), 100);
    });

    if (exportGradingBtn) {
        exportGradingBtn.addEventListener('click', () => {
            const modules = Object.keys(modulesFound);
            const data = [];
            Object.values(allStudents).forEach(student => {
                let totalScore = 0;
                modules.forEach(m => totalScore += (student.scores[m] || 0));
                const percent = (totalScore / window.totalFullScore) * 100;
                const grade = calculateGrade(percent, (window.avgMPL || 0) * 100, gradingScenarios[activeScenario]);

                data.push({
                    '#': student['#'],
                    'Username': student['username'],
                    'Firstname': student['firstname'],
                    'Surname': student['surname'],
                    'Total Score': totalScore.toFixed(2),
                    'Percentage (%)': percent.toFixed(2) + '%',
                    'Grade': grade
                });
            });
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Grading");
            
            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
            const uri = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + wbout;
            const a = document.createElement("a");
            document.body.appendChild(a);
            a.style = "display: none";
            a.href = uri;
            a.download = "Exam_Grading_Results.xlsx";
            a.click();
            setTimeout(() => document.body.removeChild(a), 100);
        });
    }
});
