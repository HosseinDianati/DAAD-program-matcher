"""
Interactive Client Portal Dashboard Generator.
Builds a standalone, responsive web application (HTML/CSS/JS) enabling:
- Real-time program browsing with color-coded Admission Chance badges.
- Dynamic Add Program (via modal dialog).
- Dynamic Remove Program (with instant re-calculation).
- Complete Required Documents collapsible checklists.
- 1-Click Export to Excel (.xlsx) generated directly in-browser using SheetJS.
- 1-Click Copy Google Sheets formatted data.
"""

import json

DASHBOARD_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Germany Study Program Advisor - Client Portal</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #0F52BA;
            --primary-dark: #0A387E;
            --secondary: #00A86B;
            --bg: #F8FAFC;
            --card-bg: #FFFFFF;
            --text-main: #0F172A;
            --text-muted: #64748B;
            --border: #E2E8F0;
            --safe-bg: #DCFCE7;
            --safe-text: #166534;
            --target-bg: #FEF9C3;
            --target-text: #854D0E;
            --reach-bg: #FEE2E2;
            --reach-text: #991B1B;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        body {
            background-color: var(--bg);
            color: var(--text-main);
            padding: 24px;
        }

        .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 16px;
            margin-bottom: 24px;
            background: var(--card-bg);
            padding: 24px 32px;
            border-radius: 16px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            border: 1px solid var(--border);
        }

        .header h1 {
            font-size: 24px;
            font-weight: 700;
            color: var(--text-main);
        }

        .header p {
            font-size: 14px;
            color: var(--text-muted);
            margin-top: 4px;
        }

        .actions-bar {
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
        }

        button {
            padding: 10px 18px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            border: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }

        .btn-primary {
            background-color: var(--primary);
            color: white;
        }

        .btn-primary:hover {
            background-color: var(--primary-dark);
        }

        .btn-success {
            background-color: var(--secondary);
            color: white;
        }

        .btn-outline {
            background: white;
            border: 1px solid var(--border);
            color: var(--text-main);
        }

        .btn-outline:hover {
            background: #F1F5F9;
        }

        .btn-danger {
            background: #EF4444;
            color: white;
            padding: 6px 12px;
            font-size: 12px;
        }

        .btn-danger:hover {
            background: #DC2626;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .stat-card {
            background: var(--card-bg);
            padding: 16px 20px;
            border-radius: 12px;
            border: 1px solid var(--border);
            display: flex;
            flex-direction: column;
            gap: 4px;
        }

        .stat-card .val {
            font-size: 24px;
            font-weight: 700;
        }

        .stat-card .lbl {
            font-size: 12px;
            color: var(--text-muted);
            text-transform: uppercase;
            font-weight: 600;
            letter-spacing: 0.5px;
        }

        .filters {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
            align-items: center;
            flex-wrap: wrap;
        }

        .filter-btn {
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 500;
            background: white;
            border: 1px solid var(--border);
            color: var(--text-muted);
            cursor: pointer;
        }

        .filter-btn.active {
            background: var(--primary);
            color: white;
            border-color: var(--primary);
        }

        .program-card {
            background: var(--card-bg);
            border-radius: 14px;
            border: 1px solid var(--border);
            margin-bottom: 20px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .program-card:hover {
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.07);
        }

        .card-header {
            padding: 18px 24px;
            background: #FFFFFF;
            border-bottom: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 16px;
        }

        .card-header .title-area h2 {
            font-size: 18px;
            font-weight: 700;
            color: var(--text-main);
        }

        .card-header .title-area .uni-name {
            font-size: 14px;
            font-weight: 500;
            color: var(--text-muted);
            margin-top: 2px;
        }

        .badge-chance {
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            white-space: nowrap;
        }

        .badge-safe { background: var(--safe-bg); color: var(--safe-text); }
        .badge-target { background: var(--target-bg); color: var(--target-text); }
        .badge-reach { background: var(--reach-bg); color: var(--reach-text); }

        .card-body {
            padding: 20px 24px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 20px;
        }

        .info-col {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }

        .info-item {
            font-size: 13px;
            line-height: 1.5;
        }

        .info-item .label {
            font-weight: 600;
            color: var(--text-muted);
            display: block;
            margin-bottom: 2px;
            font-size: 11px;
            text-transform: uppercase;
        }

        .docs-box {
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            padding: 12px 16px;
            font-size: 13px;
        }

        .docs-box h4 {
            font-size: 12px;
            color: var(--primary);
            margin-bottom: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .card-footer {
            padding: 14px 24px;
            background: #F8FAFC;
            border-top: 1px solid var(--border);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
        }

        .links-group {
            display: flex;
            gap: 10px;
        }

        .link-pill {
            text-decoration: none;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            transition: all 0.2s;
        }

        .link-prog {
            background: #E0E7FF;
            color: #3730A3;
        }

        .link-prog:hover {
            background: #C7D2FE;
        }

        .link-apply {
            background: #D1FAE5;
            color: #065F46;
        }

        .link-apply:hover {
            background: #A7F3D0;
        }

        /* Modal styling */
        .modal-overlay {
            display: none;
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(0, 0, 0, 0.5);
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }

        .modal-overlay.active {
            display: flex;
        }

        .modal-content {
            background: white;
            border-radius: 16px;
            width: 90%;
            max-width: 650px;
            max-height: 90vh;
            overflow-y: auto;
            padding: 28px;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.2);
        }

        .modal-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
        }

        .form-group {
            margin-bottom: 14px;
        }

        .form-group label {
            display: block;
            font-size: 12px;
            font-weight: 600;
            color: var(--text-muted);
            margin-bottom: 4px;
            text-transform: uppercase;
        }

        .form-control {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-size: 14px;
        }

        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }
    </style>
</head>
<body>

    <div class="header">
        <div>
            <h1>🇩🇪 Germany Higher Education - Matched Study Programs</h1>
            <p>Interactive Client Advisor Portal & Program Shortlist</p>
        </div>
        <div class="actions-bar">
            <button class="btn-primary" onclick="openAddModal()">➕ Add Program</button>
            <button class="btn-success" onclick="exportToExcel()">📥 Download Excel (.xlsx)</button>
            <button class="btn-outline" onclick="copyGoogleSheetData()">📋 Copy Google Sheet Format</button>
        </div>
    </div>

    <div class="stats-grid">
        <div class="stat-card">
            <span class="lbl">Total Programs</span>
            <span class="val" id="stat-total">0</span>
        </div>
        <div class="stat-card">
            <span class="lbl">🟢 High Chance (Safe)</span>
            <span class="val" style="color: var(--safe-text);" id="stat-safe">0</span>
        </div>
        <div class="stat-card">
            <span class="lbl">🟡 Medium Chance (Target)</span>
            <span class="val" style="color: var(--target-text);" id="stat-target">0</span>
        </div>
        <div class="stat-card">
            <span class="lbl">🔴 Low Chance (Reach)</span>
            <span class="val" style="color: var(--reach-text);" id="stat-reach">0</span>
        </div>
    </div>

    <div class="filters">
        <span style="font-size: 13px; font-weight: 600; color: var(--text-muted); margin-right: 8px;">Filter by Admission Chance:</span>
        <button class="filter-btn active" onclick="filterPrograms('all')">All Programs</button>
        <button class="filter-btn" onclick="filterPrograms('safe')">🟢 High (Safe)</button>
        <button class="filter-btn" onclick="filterPrograms('target')">🟡 Medium (Target)</button>
        <button class="filter-btn" onclick="filterPrograms('reach')">🔴 Low (Reach)</button>
    </div>

    <div id="programs-container"></div>

    <!-- Modal for Adding Programs -->
    <div class="modal-overlay" id="add-modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>Add New Study Program</h3>
                <button class="btn-outline" onclick="closeAddModal()" style="padding: 4px 10px;">✕</button>
            </div>
            <form id="add-program-form" onsubmit="handleAddProgram(event)">
                <div class="form-row">
                    <div class="form-group">
                        <label>University Name</label>
                        <input type="text" class="form-control" id="form-uni" required placeholder="e.g. RWTH Aachen University">
                    </div>
                    <div class="form-group">
                        <label>Program Name</label>
                        <input type="text" class="form-control" id="form-prog" required placeholder="e.g. Computational Engineering (M.Sc.)">
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Admission Chance</label>
                        <select class="form-control" id="form-chance">
                            <option value="High (Safe)">🟢 High (Safe)</option>
                            <option value="Medium (Target)" selected>🟡 Medium (Target)</option>
                            <option value="Low (Reach)">🔴 Low (Reach)</option>
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Tuition Fee</label>
                        <input type="text" class="form-control" id="form-tuition" value="none" placeholder="e.g. none or 1500 Euro/sem">
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Semester Fee</label>
                        <input type="text" class="form-control" id="form-semester-fee" placeholder="e.g. 320 Euro per semester">
                    </div>
                    <div class="form-group">
                        <label>Intake</label>
                        <input type="text" class="form-control" id="form-intake" value="Winter">
                    </div>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Non-EU Application Deadline</label>
                        <input type="text" class="form-control" id="form-deadline" placeholder="e.g. 15 July or 31 May">
                    </div>
                    <div class="form-group">
                        <label>Application Method</label>
                        <select class="form-control" id="form-method">
                            <option value="Uni-Assist">Uni-Assist</option>
                            <option value="Uni Portal">Uni Portal</option>
                            <option value="Uni Portal (via VPD)">Uni Portal (via VPD)</option>
                        </select>
                    </div>
                </div>

                <div class="form-group">
                    <label>Requirements & ECTS Breakdown</label>
                    <textarea class="form-control" id="form-req" rows="3" placeholder="Specific degree and credit distribution..."></textarea>
                </div>

                <div class="form-group">
                    <label>Required Documents Checklist</label>
                    <textarea class="form-control" id="form-docs" rows="3" placeholder="1. Bachelor certificate\n2. Transcript with syllabus\n3. APS certificate\n4. IELTS 7.0..."></textarea>
                </div>

                <div class="form-row">
                    <div class="form-group">
                        <label>Program Website Link</label>
                        <input type="url" class="form-control" id="form-prog-link" placeholder="https://...">
                    </div>
                    <div class="form-group">
                        <label>Apply Portal Link</label>
                        <input type="url" class="form-control" id="form-apply-link" placeholder="https://my.uni-assist.de/ or portal">
                    </div>
                </div>

                <div class="form-group">
                    <label>Special Notes</label>
                    <input type="text" class="form-control" id="form-note" placeholder="Any crucial notes or warnings...">
                </div>

                <div style="display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px;">
                    <button type="button" class="btn-outline" onclick="closeAddModal()">Cancel</button>
                    <button type="submit" class="btn-primary">Save Program</button>
                </div>
            </form>
        </div>
    </div>

    <script>
        let programs = __PROGRAMS_JSON__;
        let currentFilter = 'all';

        function render() {
            const container = document.getElementById('programs-container');
            container.innerHTML = '';

            let filtered = programs;
            if (currentFilter === 'safe') {
                filtered = programs.filter(p => p["Admission Chance"].toLowerCase().includes('high') || p["Admission Chance"].toLowerCase().includes('safe'));
            } else if (currentFilter === 'target') {
                filtered = programs.filter(p => p["Admission Chance"].toLowerCase().includes('medium') || p["Admission Chance"].toLowerCase().includes('target'));
            } else if (currentFilter === 'reach') {
                filtered = programs.filter(p => p["Admission Chance"].toLowerCase().includes('low') || p["Admission Chance"].toLowerCase().includes('reach'));
            }

            // Update stats
            document.getElementById('stat-total').innerText = programs.length;
            document.getElementById('stat-safe').innerText = programs.filter(p => p["Admission Chance"].toLowerCase().includes('high') || p["Admission Chance"].toLowerCase().includes('safe')).length;
            document.getElementById('stat-target').innerText = programs.filter(p => p["Admission Chance"].toLowerCase().includes('medium') || p["Admission Chance"].toLowerCase().includes('target')).length;
            document.getElementById('stat-reach').innerText = programs.filter(p => p["Admission Chance"].toLowerCase().includes('low') || p["Admission Chance"].toLowerCase().includes('reach')).length;

            if (filtered.length === 0) {
                container.innerHTML = `<div style="text-align: center; padding: 40px; color: var(--text-muted); background: white; border-radius: 12px; border: 1px solid var(--border);">No programs match the selected filter.</div>`;
                return;
            }

            filtered.forEach((p, idx) => {
                const card = document.createElement('div');
                card.className = 'program-card';

                let badgeClass = 'badge-target';
                const chanceLower = p["Admission Chance"].toLowerCase();
                if (chanceLower.includes('high') || chanceLower.includes('safe')) badgeClass = 'badge-safe';
                if (chanceLower.includes('low') || chanceLower.includes('reach')) badgeClass = 'badge-reach';

                card.innerHTML = `
                    <div class="card-header">
                        <div class="title-area">
                            <h2>${p["Program"]}</h2>
                            <div class="uni-name">🏛️ ${p["University"]}</div>
                        </div>
                        <div style="display: flex; gap: 8px; align-items: center;">
                            <span class="badge-chance ${badgeClass}">${p["Admission Chance"]}</span>
                            <button class="btn-danger" onclick="removeProgram(${idx})">🗑️ Remove</button>
                        </div>
                    </div>
                    <div class="card-body">
                        <div class="info-col">
                            <div class="info-item">
                                <span class="label">📅 Application Deadline (Non-EU)</span>
                                <strong>${p["Deadline"]}</strong>
                            </div>
                            <div class="info-item">
                                <span class="label">💰 Tuition & Semester Fees</span>
                                <span>Tuition: <strong>${p["Tuition fee"]}</strong> | Semester Fee: <strong>${p["Semester fee"]}</strong></span>
                            </div>
                            <div class="info-item">
                                <span class="label">🚀 Intake & Application Method</span>
                                <span>${p["Intake"]} Intake via <strong>${p["Application Method"]}</strong></span>
                            </div>
                            ${p["Note"] ? `
                            <div class="info-item" style="background: #F1F5F9; padding: 8px 12px; border-radius: 6px;">
                                <span class="label">💡 Note / Remarks</span>
                                <span>${p["Note"]}</span>
                            </div>` : ''}
                        </div>
                        <div class="info-col">
                            <div class="docs-box">
                                <h4>📋 Academic Requirements & ECTS</h4>
                                <p style="white-space: pre-line;">${p["Requirements"]}</p>
                            </div>
                            <div class="docs-box" style="margin-top: 10px; background: #FAF5FF; border-color: #E9D5FF;">
                                <h4 style="color: #6B21A8;">📂 Required Documents Checklist</h4>
                                <p style="white-space: pre-line;">${p["Required Documents"] || "Standard German application documents (Transcript, Degree Certificate, APS, Language proof)."}</p>
                            </div>
                        </div>
                    </div>
                    <div class="card-footer">
                        <div style="font-size: 12px; color: var(--text-muted);">
                            Apply via: <strong>${p["Application Method"]}</strong>
                        </div>
                        <div class="links-group">
                            ${p["Program link"] ? `<a href="${p["Program link"]}" target="_blank" class="link-pill link-prog">🌐 Official Program Page</a>` : ''}
                            ${p["Apply Link"] ? `<a href="${p["Apply Link"]}" target="_blank" class="link-pill link-apply">✍️ Direct Apply Portal</a>` : ''}
                        </div>
                    </div>
                `;
                container.appendChild(card);
            });
        }

        function filterPrograms(type) {
            currentFilter = type;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            event.target.classList.add('active');
            render();
        }

        function removeProgram(index) {
            if (confirm(`Are you sure you want to remove "${programs[index]["Program"]}" from the list?`)) {
                programs.splice(index, 1);
                render();
            }
        }

        function openAddModal() {
            document.getElementById('add-modal').classList.add('active');
        }

        function closeAddModal() {
            document.getElementById('add-modal').classList.remove('active');
        }

        function handleAddProgram(e) {
            e.preventDefault();
            const newProg = {
                "University": document.getElementById('form-uni').value,
                "Program": document.getElementById('form-prog').value,
                "Admission Chance": document.getElementById('form-chance').value,
                "Tuition fee": document.getElementById('form-tuition').value,
                "Semester fee": document.getElementById('form-semester-fee').value,
                "Intake": document.getElementById('form-intake').value,
                "Deadline": document.getElementById('form-deadline').value,
                "Requirements": document.getElementById('form-req').value,
                "Required Documents": document.getElementById('form-docs').value,
                "Application Method": document.getElementById('form-method').value,
                "Program link": document.getElementById('form-prog-link').value,
                "Note": document.getElementById('form-note').value,
                "Apply Link": document.getElementById('form-apply-link').value
            };
            programs.unshift(newProg);
            closeAddModal();
            document.getElementById('add-program-form').reset();
            render();
        }

        function exportToExcel() {
            const headers = [
                "University", "Program", "Admission Chance", "Tuition fee", "Semester fee",
                "Intake", "Deadline", "Requirements", "Required Documents",
                "Application Method", "Program link", "Note", "Apply Link"
            ];
            const data = [headers];
            programs.forEach(p => {
                data.push(headers.map(h => p[h] || ""));
            });

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(wb, ws, "Germany Programs");
            XLSX.writeFile(wb, "German_University_Matched_Programs.xlsx");
        }

        function copyGoogleSheetData() {
            const headers = [
                "University", "Program", "Admission Chance", "Tuition fee", "Semester fee",
                "Intake", "Deadline", "Requirements", "Required Documents",
                "Application Method", "Program link", "Note", "Apply Link"
            ];
            let tsv = headers.join("\t") + "\n";
            programs.forEach(p => {
                const row = headers.map(h => (p[h] || "").replace(/\t/g, " ").replace(/\n/g, " "));
                tsv += row.join("\t") + "\n";
            });

            navigator.clipboard.writeText(tsv).then(() => {
                alert("Copied data in Google Sheets TSV format! Simply paste (Ctrl+V) into any Google Sheet.");
            }).catch(err => {
                alert("Failed to copy automatically: " + err);
            });
        }

        render();
    </script>
</body>
</html>
"""

def generate_client_dashboard(rows: list[dict], output_html_path: str):
    """
    Injects program data into the dashboard template and writes the HTML file.
    """
    programs_json_str = json.dumps(rows, ensure_ascii=False, indent=2)
    html_content = DASHBOARD_TEMPLATE.replace("__PROGRAMS_JSON__", programs_json_str)
    
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Generated interactive client dashboard: {output_html_path}")
