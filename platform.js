document.addEventListener('DOMContentLoaded', () => {
    // Initialize Lucide icons
    lucide.createIcons();

    // --- NAVIGATION LOGIC ---
    const navItems = document.querySelectorAll('.nav-item');
    const toolSections = document.querySelectorAll('.tool-section');

    navItems.forEach(item => {
        item.addEventListener('click', () => {
            const target = item.dataset.target;
            
            // Update Nav
            navItems.forEach(i => i.classList.remove('active'));
            item.classList.add('active');

            // Update Sections
            toolSections.forEach(sec => {
                sec.classList.remove('active');
                if (sec.id === target) sec.classList.add('active');
            });
        });
    });

    // =========================================================================
    // --- AGGREGATOR LOGIC (from app.js) ---
    // =========================================================================
    
    // Selectors
    const aggDropZone = document.getElementById('agg-drop-zone');
    const aggFileInput = document.getElementById('agg-file-input');
    const aggFileListEl = document.getElementById('agg-file-list');
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
    const gradingSection = document.getElementById('grading-section');
    const gradingHeader = document.getElementById('grading-header');
    const gradingBody = document.getElementById('grading-body');
    const exportScoresBtn = document.getElementById('export-scores-btn');
    const exportFeedbackBtn = document.getElementById('export-feedback-btn');
    const exportGradingBtn = document.getElementById('export-grading-btn');
    const exportStatsBtn = document.getElementById('export-stats-btn');
    const downloadChartBtn = document.getElementById('download-chart-btn');

    let allStudents = {}; 
    let modulesFound = {}; 
    let incrementalRate = 4;
    let gradingScenarios = { 1: 4.0, 2: 3.0, 3: 3.5 };
    let activeScenario = 1;
    let currentSort = { column: '#', direction: 'asc' };

    // Handlers
    if (aggDropZone) {
        aggDropZone.addEventListener('click', () => aggFileInput.click());
        aggDropZone.addEventListener('dragover', (e) => { e.preventDefault(); aggDropZone.classList.add('dragover'); });
        aggDropZone.addEventListener('dragleave', () => aggDropZone.classList.remove('dragover'));
        aggDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            aggDropZone.classList.remove('dragover');
            handleAggFiles(e.dataTransfer.files);
        });
        aggFileInput.addEventListener('change', (e) => handleAggFiles(e.target.files));
    }

    async function handleAggFiles(files) {
        if (files.length === 0) return;
        const uploadedFiles = Array.from(files);
        aggFileListEl.innerHTML = '';
        allStudents = {};
        modulesFound = {};

        for (const file of uploadedFiles) {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `<span><i data-lucide="file-text"></i> ${file.name}</span>`;
            aggFileListEl.appendChild(item);
            await processAggFile(file);
        }
        lucide.createIcons();
        updateAggregates();
    }

    async function processAggFile(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);
                
                jsonData.forEach(row => {
                    const username = row['username'] || row['Username'];
                    if (!username) return;

                    if (!allStudents[username]) {
                        allStudents[username] = {
                            '#': row['#'] || '',
                            'username': username,
                            'firstname': row['firstname'] || row['Firstname'] || '',
                            'surname': row['surname'] || row['Surname'] || '',
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
                                    mpl: 0.44 // Default
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

    function updateAggregates() {
        const modules = Object.keys(modulesFound);
        if (modules.length === 0) return;

        configSection.classList.remove('hidden');
        moduleConfigGrid.innerHTML = '';
        
        let totalFullScore = 0;
        let weightedMPLSum = 0;

        modules.forEach((modName, index) => {
            const mod = modulesFound[modName];
            const colorVar = `--mod-${(index % 4) + 1}`;
            
            const card = document.createElement('div');
            card.className = 'module-card';
            card.style.borderLeft = `4px solid var(${colorVar})`;
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
                    Calculated Pass: <span class="raw">${(mod.fullScore * mod.mpl).toFixed(2)}</span> | Ceil: <span class="ceil">${Math.ceil(mod.fullScore * mod.mpl)}</span>
                </div>
            `;
            moduleConfigGrid.appendChild(card);
            
            totalFullScore += mod.fullScore;
            weightedMPLSum += (mod.fullScore * mod.mpl);
        });

        window.totalFullScore = totalFullScore;
        const avgMPL = totalFullScore > 0 ? (weightedMPLSum / totalFullScore) : 0;
        window.avgMPL = avgMPL;
        const totalPassMark = weightedMPLSum;
        const totalPassMarkCeil = Math.ceil(totalPassMark);
        window.totalPassMarkCeil = totalPassMarkCeil;

        document.getElementById('stat-count').textContent = `${modules.length} / 4`;
        document.getElementById('stat-avg-mpl').textContent = (avgMPL * 100).toFixed(2) + '%';
        document.getElementById('stat-total-pass').textContent = `${totalPassMark.toFixed(2)} (Ceil: ${totalPassMarkCeil})`;

        // Re-bind inputs
        moduleConfigGrid.querySelectorAll('input').forEach(input => {
            input.addEventListener('input', (e) => {
                const mod = e.target.dataset.mod;
                const type = e.target.dataset.type;
                modulesFound[mod][type] = parseFloat(e.target.value) || 0;
                updateAggregates();
            });
        });

        renderResultsTable();
        renderFeedbackTable();
        renderGradingSummaries();
        renderGradingTable();
        renderBoxPlot();
        renderStatsTable();
    }

    // Incremental Rate listener
    const irInput = document.getElementById('incremental-rate');
    if (irInput) {
        irInput.addEventListener('input', (e) => {
            incrementalRate = parseFloat(e.target.value) || 0;
            updateAggregates();
        });
    }

    // Grading listeners
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
            document.querySelectorAll('.apply-scenario-btn').forEach(b => b.classList.add('secondary-style'));
            btn.classList.remove('secondary-style');
            renderGradingTable();
        });
    });

    // Rendering functions (Clean versions)
    function calculateGrade(percent, mplPercent, step) {
        if (percent < mplPercent) return 'D/F';
        if (percent < mplPercent + step) return 'D+';
        if (percent < mplPercent + 2 * step) return 'C';
        if (percent < mplPercent + 3 * step) return 'C+';
        if (percent < mplPercent + 4 * step) return 'B';
        if (percent < mplPercent + 5 * step) return 'B+';
        return 'A';
    }

    function getFeedbackCategory(score, fullScore, mpl) {
        if (!fullScore || fullScore === 0) return { label: 'N/A', class: '' };
        const scorePercent = (score / fullScore) * 100;
        const mplPercent = mpl * 100;
        if (scorePercent >= mplPercent + 20) return { label: 'Very good', class: 'fb-very-good' };
        if (scorePercent >= mplPercent + 12) return { label: 'Good', class: 'fb-good' };
        if (scorePercent >= mplPercent + 4) return { label: 'Borderline', class: 'fb-borderline' };
        return { label: 'Needs improvement', class: 'fb-needs-improvement' };
    }

    function calculateStats(scores) {
        if (scores.length === 0) return { min: 0, max: 0, mean: 0, sd: 0, median: 0, q1: 0, q3: 0, iqr: 0 };
        const sorted = [...scores].sort((a, b) => a - b);
        const mean = sorted.reduce((a, b) => a + b, 0) / sorted.length;
        const sd = Math.sqrt(sorted.map(x => Math.pow(x - mean, 2)).reduce((a, b) => a + b, 0) / sorted.length);
        return { 
            min: sorted[0], max: sorted[sorted.length-1], mean, sd, 
            median: sorted[Math.floor(sorted.length/2)], 
            q1: sorted[Math.floor(sorted.length/4)], 
            q3: sorted[Math.floor(sorted.length*0.75)],
            iqr: sorted[Math.floor(sorted.length*0.75)] - sorted[Math.floor(sorted.length/4)]
        };
    }

    function renderResultsTable() {
        const modules = Object.keys(modulesFound);
        resultsSection.classList.remove('hidden');
        let headerHtml = `<th class="sortable" data-col="#">#</th><th>Username</th><th>Firstname</th><th>Surname</th>`;
        modules.forEach(mod => headerHtml += `<th class="sortable" data-col="${mod}">${mod}</th>`);
        headerHtml += `<th class="sortable" data-col="sum">SUM (${window.totalFullScore})</th>`;
        tableHeader.innerHTML = headerHtml;

        const sortedStudents = Object.values(allStudents).sort((a, b) => {
            let valA, valB;
            if (currentSort.column === '#') {
                valA = parseInt(a['#']) || 0;
                valB = parseInt(b['#']) || 0;
            } else if (currentSort.column === 'sum') {
                valA = Object.values(a.scores).reduce((s,v)=>s+v,0);
                valB = Object.values(b.scores).reduce((s,v)=>s+v,0);
            } else {
                valA = a.scores[currentSort.column] || 0;
                valB = b.scores[currentSort.column] || 0;
            }
            return currentSort.direction === 'asc' ? valA - valB : valB - valA;
        });

        tableBody.innerHTML = '';
        sortedStudents.forEach(student => {
            const tr = document.createElement('tr');
            let rowSum = 0;
            let rowHtml = `<td>${student['#']}</td><td>${student['username']}</td><td>${student['firstname']}</td><td>${student['surname']}</td>`;
            modules.forEach((modName, idx) => {
                const score = student.scores[modName] || 0;
                rowSum += score;
                const isFail = score < (modulesFound[modName].fullScore * modulesFound[modName].mpl);
                rowHtml += `<td class="${isFail ? 'fail-mod-'+((idx%4)+1) : ''}">${score}</td>`;
            });
            rowHtml += `<td class="${rowSum < window.totalPassMarkCeil ? 'fail-sum' : ''}" style="font-weight:bold;">${rowSum.toFixed(2)}</td>`;
            tr.innerHTML = rowHtml;
            tableBody.appendChild(tr);
        });

        // Add Sort Listeners
        document.querySelectorAll('th.sortable').forEach(th => {
            th.onclick = () => {
                const col = th.dataset.col;
                if (currentSort.column === col) {
                    currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
                } else {
                    currentSort.column = col;
                    currentSort.direction = 'asc';
                }
                renderResultsTable();
            };
            
            // Add sort indicator
            if (th.dataset.col === currentSort.column) {
                const baseText = th.textContent.replace(/[▴▾]/g, '').trim();
                th.innerHTML = baseText + (currentSort.direction === 'asc' ? ' ▴' : ' ▾');
                th.style.color = 'var(--primary)';
            }
        });
    }

    function renderFeedbackTable() {
        const modules = Object.keys(modulesFound);
        feedbackSection.classList.remove('hidden');
        feedbackHeader.innerHTML = `<th>#</th><th>Username</th><th>Firstname</th><th>Surname</th>` + modules.map(m => `<th>${m} Feedback</th>`).join('');
        feedbackBody.innerHTML = '';
        Object.values(allStudents).forEach(student => {
            const tr = document.createElement('tr');
            let rowHtml = `<td>${student['#']}</td><td>${student['username']}</td><td>${student['firstname']}</td><td>${student['surname']}</td>`;
            modules.forEach(modName => {
                const fb = getFeedbackCategory(student.scores[modName]||0, modulesFound[modName].fullScore, modulesFound[modName].mpl);
                rowHtml += `<td><span class="feedback-badge ${fb.class}">${fb.label}</span></td>`;
            });
            tr.innerHTML = rowHtml;
            feedbackBody.appendChild(tr);
        });
    }

    function renderGradingSummaries() {
        const grades = ['D/F', 'D+', 'C', 'C+', 'B', 'B+', 'A'];
        const mplPercent = (window.avgMPL || 0) * 100;
        [1, 2, 3].forEach(id => {
            const step = gradingScenarios[id];
            const counts = { 'D/F': 0, 'D+': 0, 'C': 0, 'C+': 0, 'B': 0, 'B+': 0, 'A': 0 };
            Object.values(allStudents).forEach(s => {
                let t = Object.values(s.scores).reduce((a,b)=>a+b,0);
                counts[calculateGrade((t/window.totalFullScore)*100, mplPercent, step)]++;
            });
            const body = document.getElementById(`summary-body-${id}`);
            if (!body) return;
            let html = `<tr><td>Min % Score</td>${grades.map((g,i)=>`<td>${i===0?'0.00':(i===1?mplPercent:(mplPercent+(i-1)*step)).toFixed(2)}%</td>`).join('')}</tr>`;
            html += `<tr><td>Min Points</td>${grades.map((g,i)=>`<td>${i===0?'0.00':(window.totalFullScore*(i===1?mplPercent:(mplPercent+(i-1)*step))/100).toFixed(2)}</td>`).join('')}</tr>`;
            html += `<tr><td>No. Students</td>${grades.map(g=>`<td>${counts[g]}</td>`).join('')}</tr>`;
            body.innerHTML = html;
        });
    }

    function renderGradingTable() {
        gradingSection.classList.remove('hidden');
        gradingHeader.innerHTML = `<th>#</th><th>Username</th><th>Firstname</th><th>Surname</th><th>Total Score</th><th>%</th><th>Grade</th>`;
        gradingBody.innerHTML = '';
        const step = gradingScenarios[activeScenario];
        const mplP = (window.avgMPL || 0) * 100;
        Object.values(allStudents).forEach(student => {
            const total = Object.values(student.scores).reduce((a,b)=>a+b,0);
            const p = (total / window.totalFullScore) * 100;
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${student['#']}</td><td>${student['username']}</td><td>${student['firstname']}</td><td>${student['surname']}</td><td>${total.toFixed(2)}</td><td>${p.toFixed(2)}%</td><td><div class="grade-cell-highlight">${calculateGrade(p, mplP, step)}</div></td>`;
            gradingBody.appendChild(tr);
        });
    }

    function renderBoxPlot() {
        chartSection.classList.remove('hidden');
        document.getElementById('performance-chart-container').classList.add('performance-chart-white');
        
        const modules = Object.keys(modulesFound);
        const plotData = [];
        const colors = ["#D62728", "#1F77B4", "#2CA02C", "#9467BD", "#FF7F0E", "#222222"];

        // Individual Modules
        modules.forEach((modName, idx) => {
            const mod = modulesFound[modName];
            const scores = Object.values(allStudents).map(s => ((s.scores[modName] || 0) / mod.fullScore) * 100);
            const mplPercent = mod.mpl * 100;
            
            plotData.push({
                y: scores,
                type: 'box',
                name: modName,
                boxpoints: 'all',
                jitter: 0.5,
                pointpos: 0,
                marker: { color: colors[idx % colors.length], size: 4, opacity: 0.5 },
                fillcolor: colors[idx % colors.length],
                line: { color: colors[idx % colors.length], width: 1 },
                opacity: 0.8
            });

            // Add MPL Marker 'x'
            plotData.push({
                x: [modName],
                y: [mplPercent],
                mode: 'markers',
                name: modName + ' MPL',
                marker: { symbol: 'x', size: 12, color: '#ff0000', line: { width: 2 } },
                showlegend: false
            });
        });

        // Overall Score Box
        const overallScores = Object.values(allStudents).map(s => {
            const total = Object.values(s.scores).reduce((a,b)=>a+b,0);
            return (total / window.totalFullScore) * 100;
        });
        const overallMPL = (window.avgMPL || 0) * 100;

        plotData.push({
            y: overallScores,
            type: 'box',
            name: 'OVERALL',
            boxpoints: 'all',
            jitter: 0.5,
            pointpos: 0,
            marker: { color: '#333333', size: 5, opacity: 0.6 },
            fillcolor: '#eeeeee',
            line: { color: '#000000', width: 2 }
        });

        plotData.push({
            x: ['OVERALL'],
            y: [overallMPL],
            mode: 'markers',
            name: 'Overall MPL',
            marker: { symbol: 'x', size: 14, color: '#ff0000', line: { width: 3 } },
            showlegend: false
        });

        const layout = {
            title: { text: 'Exam Performance Distribution', font: { size: 20, color: '#0f172a' } },
            paper_bgcolor: 'white',
            plot_bgcolor: 'white',
            font: { color: '#1e293b', family: 'Outfit' },
            yaxis: { title: 'Score (%)', range: [0, 105], gridcolor: '#f1f5f9', zerolinecolor: '#e2e8f0' },
            xaxis: { gridcolor: 'transparent' },
            margin: { t: 80, b: 100, l: 80, r: 40 },
            showlegend: true,
            legend: { 
                orientation: 'h', 
                y: -0.2, 
                x: 0.5, 
                xanchor: 'center',
                font: { size: 12 }
            },
            dragmode: false,
            autosize: true
        };

        Plotly.newPlot('performance-chart-container', plotData, layout, { responsive: true, displayModeBar: false });
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
            
            modules.forEach(modName => {
                const scores = Object.values(allStudents).map(s => s.scores[modName] || 0);
                const s = calculateStats(scores);
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

    // Export Handlers
    exportScoresBtn.onclick = () => {
        const modules = Object.keys(modulesFound);
        const data = Object.values(allStudents).map(s => {
            const row = { '#': s['#'], Username: s.username, Firstname: s.firstname, Surname: s.surname };
            modules.forEach(m => row[m] = s.scores[m] || 0);
            row['Total'] = Object.values(s.scores).reduce((a,b)=>a+b,0);
            return row;
        });
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Scores");
        XLSX.writeFile(wb, "Exam_Scores.xlsx");
    };

    exportFeedbackBtn.onclick = () => {
        const modules = Object.keys(modulesFound);
        const data = Object.values(allStudents).map(s => {
            const row = { Username: s.username, Firstname: s.firstname, Surname: s.surname };
            modules.forEach(m => {
                const fb = getFeedbackCategory(s.scores[m]||0, modulesFound[m].fullScore, modulesFound[m].mpl);
                row[m + ' Feedback'] = fb.label;
            });
            return row;
        });
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Feedback");
        XLSX.writeFile(wb, "Exam_Feedback.xlsx");
    };

    exportGradingBtn.onclick = () => {
        const modules = Object.keys(modulesFound);
        const step = gradingScenarios[activeScenario];
        const mplP = (window.avgMPL || 0) * 100;
        const data = Object.values(allStudents).map(s => {
            const total = Object.values(s.scores).reduce((a,b)=>a+b,0);
            const p = (total / window.totalFullScore) * 100;
            return {
                '#': s['#'], Username: s.username, Firstname: s.firstname, Surname: s.surname,
                'Total Score': total.toFixed(2), 'Percentage (%)': p.toFixed(2) + '%',
                'Grade': calculateGrade(p, mplP, step)
            };
        });
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Grading");
        XLSX.writeFile(wb, "Exam_Grading.xlsx");
    };

    downloadChartBtn.onclick = () => {
        Plotly.downloadImage('performance-chart-container', { format: 'png', filename: 'Exam_Performance_Chart' });
    };

    if (exportStatsBtn) {
        exportStatsBtn.onclick = () => {
            const modules = Object.keys(modulesFound);
            const metrics = [
                { label: 'Max', key: 'max' },
                { label: 'Min', key: 'min' },
                { label: 'Mean', key: 'mean' },
                { label: 'SD', key: 'sd' },
                { label: 'Median', key: 'median' },
                { label: 'IQR', key: 'iqr' }
            ];
            const data = metrics.map(m => {
                const row = { Metric: m.label };
                modules.forEach(modName => {
                    const s = calculateStats(Object.values(allStudents).map(st => st.scores[modName] || 0));
                    row[modName] = s[m.key].toFixed(2);
                });
                // Add Total Stats
                const totalScores = Object.values(allStudents).map(student => {
                    return Object.keys(modulesFound).reduce((sum, mod) => sum + (student.scores[mod] || 0), 0);
                });
                const ts = calculateStats(totalScores);
                row['TOTAL'] = ts[m.key].toFixed(2);
                return row;
            });
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Statistics");
            XLSX.writeFile(wb, "Exam_Statistics.xlsx");
        };
    }

    // =========================================================================
    // --- PARSER LOGIC (from parser.js) ---
    // =========================================================================
    const GEMINI_API_KEY = "AIzaSyBn-EN72b2iUdbgP15ecGoGl9OjyuHeEI8"; 
    const SYSTEM_PROMPT = `You are an advanced medical exam document parser and formatter for the Exam+ system. You will receive raw text or Word document content containing multiple-choice questions.

### CRITICAL RULE:
**STRICTLY PRESERVE THE ORIGINAL TEXT.** Do not paraphrase, summarize, or rewrite the medical content, clinical vignettes, or choices. Your job is ONLY to restructure and format, NOT to edit the language.

Your task is to process the input in TWO phases.

### PHASE 1: TEXT RECOGNITION & STRUCTURING (For Raw/Unformatted Input)
If the input text lacks the standard $VAR tags, you must first extract and structure each question into the exact following order: $TOPIC, $Q, $A, $B, $C, $D, $E, $ANSWER, $tags.

Apply these parsing rules:
1. **Identify $Q$ (Question Stem):** Locate the main body of the question (often a clinical vignette). **Remove the original question numbers** (e.g., "1.", "ข้อ 1"). **DO NOT CHANGE ANY WORDS.** Preserve all lab result lines, tables, or line breaks within the vignette exactly as they appear.
2. **Identify Choices ($A to $E):** Locate the options. Map them exactly to $A, $B, $C, $D, and $E. Maintain the original wording of every choice.
3. **Identify $ANSWER and $tags:** Extract the correct answer and any metadata/keywords provided into their respective tags. If $TOPIC is not explicitly provided, use $TOPIC for now.
4. **Spacing Rule:** There MUST be exactly **one blank line** between every different $ tag.

### PHASE 2: EXAM+ FORMATTING RULES
Once structured, apply these layout rules:
1. **CONSOLIDATE TAGS:** Move any original $TOPIC and $SUBTOPIC values into $tags, appended as comma-separated values.
2. **SET NEW $TOPIC:** Replace all $TOPIC lines with: $TOPIC {yyyy}_class{xxxx}
3. **EMPHASIZE NEGATIVE KEYWORDS:** Scan the lead-in question (the final sentence of $Q$) for "not", "least", or "except". You MUST format these specific words as UPPERCASE AND BOLD (i.e., **NOT**, **LEAST**, **EXCEPT**).
4. **CHOICE PUNCTUATION ($A - $E):** 
   - If the choice is a complete sentence, it MUST end with a full stop (.).
   - If the choice is a word or fragment, do NOT use a full stop.
5. **ONE QUESTION PER PAGE:** Each block must occupy exactly one page (handled by the system, but ensure output is clear).
6. **HEADER LOGIC:** (Handled by system)
7. **FONT:** (Handled by system)

### IMPORTANT — IMAGES AND TABLES:
If a question contains an image or complex table that cannot be represented in text, include $MANUAL: [description] at the end of that question block.

Output the results with each question separated by "---PAGE_BREAK---".`;

    const parserDropZone = document.getElementById('parser-drop-zone');
    const parserFileInput = document.getElementById('parser-file-input');
    const parserProcessBtn = document.getElementById('parser-process-btn');
    const parserDownloadBtn = document.getElementById('parser-download-btn');
    const parserStatusText = document.getElementById('parser-status-text');
    const parserProgressBar = document.getElementById('parser-progress-bar');
    const parserReviewItems = document.getElementById('parser-review-items');
    const parserManualReview = document.getElementById('parser-manual-review');
    const parserModuleInput = document.getElementById('parser-module-name');
    const parserClassInput = document.getElementById('parser-class-year');

    let parserFile = null;
    let parsedQuestions = [];

    if (parserDropZone) {
        parserDropZone.onclick = () => parserFileInput.click();
        parserDropZone.addEventListener('dragover', (e) => { e.preventDefault(); parserDropZone.classList.add('dragover'); });
        parserDropZone.addEventListener('dragleave', () => parserDropZone.classList.remove('dragover'));
        parserDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            parserDropZone.classList.remove('dragover');
            if (e.dataTransfer.files && e.dataTransfer.files[0]) handleParserFile(e.dataTransfer.files[0]);
        });
        parserFileInput.onchange = (e) => {
            if (e.target.files && e.target.files[0]) handleParserFile(e.target.files[0]);
        };
    }

    function handleParserFile(file) {
        parserFile = file;
        document.getElementById('parser-file-name').textContent = parserFile.name;
        document.getElementById('parser-file-info').classList.remove('hidden');
        parserDropZone.classList.add('hidden');
        parserProcessBtn.disabled = false;
    }

    async function listAvailableModels(apiKey) {
        const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey.trim()}`;
        try {
            const response = await fetch(url);
            const data = await response.json();
            if (data.models) return data.models.map(m => m.name);
            return [];
        } catch (e) { return []; }
    }

    async function callGemini(apiKey, systemPrompt, userContent, modelName) {
        // modelName comes from listAvailableModels as 'models/gemini-1.5-flash'
        const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent`;
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey.trim() },
            body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: systemPrompt + "\n\nINPUT CONTENT:\n" + userContent }] }] })
        });
        const data = await response.json();
        if (data.error) throw new Error(data.error.message);
        if (!data.candidates || data.candidates.length === 0) throw new Error("AI did not return results.");
        return data.candidates[0].content.parts[0].text;
    }

    parserProcessBtn.onclick = async () => {
        if (!parserFile) return;
        const mod = parserModuleInput.value;
        const year = parserClassInput.value;
        document.getElementById('parser-progress-section').classList.remove('hidden');
        
        try {
            const ab = await parserFile.arrayBuffer();
            const { value: rawText } = await mammoth.extractRawText({ arrayBuffer: ab });
            
            parserStatusText.textContent = "Checking API Models...";
            parserProgressBar.style.width = "20%";

            const availableModels = await listAvailableModels(GEMINI_API_KEY);
            let bestModel = "models/gemini-1.5-flash";
            if (availableModels.length > 0) {
                if (availableModels.includes("models/gemini-1.5-flash")) bestModel = "models/gemini-1.5-flash";
                else if (availableModels.includes("models/gemini-1.5-flash-latest")) bestModel = "models/gemini-1.5-flash-latest";
                else if (availableModels.includes("models/gemini-1.5-pro")) bestModel = "models/gemini-1.5-pro";
                else bestModel = availableModels[0];
            }

            parserStatusText.textContent = `Using model: ${bestModel.split('/')[1]}...`;
            parserProgressBar.style.width = "50%";

            const finalPrompt = SYSTEM_PROMPT.replace(/{yyyy}/g, mod).replace(/{xxxx}/g, year);
            const aiText = await callGemini(GEMINI_API_KEY, finalPrompt, rawText, bestModel);

            parserStatusText.textContent = "Parsing Response...";
            parserProgressBar.style.width = "80%";

            parsedQuestions = aiText.split('---PAGE_BREAK---').map(q => {
                const lines = q.trim().split('\n');
                const obj = {}; let currentTag = '';
                lines.forEach(line => {
                    if (line.startsWith('$')) {
                        const spaceIndex = line.indexOf(' ');
                        const rawTag = spaceIndex !== -1 ? line.substring(0, spaceIndex) : line;
                        currentTag = rawTag.endsWith(':') ? rawTag.slice(0, -1) : rawTag;
                        obj[currentTag] = (obj[currentTag] || '') + line.substring(rawTag.length).trim();
                    } else if (currentTag) obj[currentTag] += '\n' + line;
                });
                return obj;
            });

            parserProgressBar.style.width = "100%";
            parserStatusText.textContent = "Complete!";
            document.getElementById('parser-results-section').classList.remove('hidden');

            const manualItems = parsedQuestions
                .map((q, i) => ({ ...q, originalIndex: i + 1 }))
                .filter(q => q['$MANUAL']);

            if (manualItems.length > 0) {
                parserManualReview.classList.remove('hidden');
                parserReviewItems.innerHTML = manualItems.map((m) => `<li><strong>ข้อที่ ${m.originalIndex}:</strong> ${m['$MANUAL']}</li>`).join('');
            }
        } catch (err) {
            alert("Error: " + err.message);
            parserStatusText.textContent = "Failed: " + err.message;
        }
    };

    parserDownloadBtn.onclick = async () => {
        const docxLib = window.docx || (typeof docx !== 'undefined' ? docx : null);
        if (!docxLib) return alert("Word library not loaded");
        const { Document, Packer, Paragraph, TextRun, Header } = docxLib;
        
        const sections = parsedQuestions.map((q, i) => ({
            properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
            headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: `Question ${i+1}`, size: 24 })] })] }) },
            children: Object.entries(q).map(([tag, val]) => new Paragraph({ 
                children: [new TextRun({ text: `${tag} `, bold: true, size: 24 }), new TextRun({ text: val, size: 24 })], 
                spacing: { after: 200 } 
            }))
        }));

        const doc = new Document({ sections });
        const blob = await Packer.toBlob(doc);
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Formatted_Exam_${parserModuleInput.value || 'Results'}.docx`;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { document.body.removeChild(a); window.URL.revokeObjectURL(url); }, 100);
    };

});
