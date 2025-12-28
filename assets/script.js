/* script.js - Final Version (Support for Tab, Colon, and Equals) */

// --- GLOBAL VARIABLES ---
let projects = [];
let activeProjectIdx = 0;
let isOverviewMode = false;
const MAX_TOTAL_SETS = 50; 

// Undo History Variables
let deletedItem = null;      // Stores the last deleted Set
let deletedItemIdx = -1;     // Stores where it was
let deletedRowsHistory = []; // Stores stack of deleted rows for active project

// --- INITIALIZATION ---
window.onload = function() {
    addNewProject();
};

function updateAddButtonText() {
    const input = document.getElementById('addSetQty');
    const btn = document.getElementById('btnAddSets');
    if (input && btn) {
        let qty = parseInt(input.value) || 0;
        if(qty < 1) qty = 1;
        btn.innerText = `+ Add ${qty} Set${qty > 1 ? 's' : ''}`;
    }
}

function setMode(side, mode) {
    const selectId = `headerMode${side}`;
    document.getElementById(selectId).value = mode;
    const group = document.getElementById(`groupMode${side}`);
    const buttons = group.querySelectorAll('.toggle-btn');
    buttons.forEach(btn => {
        if(btn.getAttribute('data-val') === mode) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

// --- PROJECT MANAGEMENT (SETS) ---
function addSetsFromInput() {
    const input = document.getElementById('addSetQty');
    let qty = parseInt(input.value) || 1;
    if (qty < 1) qty = 1;
    if (projects.length + qty > MAX_TOTAL_SETS) {
        showModal("Limit Reached", `Cannot add more than ${MAX_TOTAL_SETS} sets.`, 'error');
        return;
    }
    for(let i=0; i<qty; i++) createSet();
    renderTopBar();
    switchProject(projects.length - qty);
    input.value = 1;
    updateAddButtonText();
}

function addNewProject() {
    if (projects.length >= MAX_TOTAL_SETS) return;
    createSet();
    renderTopBar();
    switchProject(projects.length - 1);
}

function createSet() {
    const id = projects.length + 1;
    projects.push({
        name: `Set ${id}`,
        status: 'empty', 
        rawA: "", rawB: "", rawMatrix: "",
        dataA: null, dataB: null, matrix: [],
        mapping: [], step: 1, showMatrix: false,
        summary: { matches: 0, mismatches: 0 }
    });
}

// --- UNDO SET LOGIC ---
function deleteProject(e, idx) {
    e.stopPropagation(); 
    // Save for Undo
    deletedItem = projects[idx];
    deletedItemIdx = idx;
    
    projects.splice(idx, 1);
    
    if (projects.length === 0) {
        createSet();
        activeProjectIdx = 0;
    } else {
        if (activeProjectIdx >= projects.length) activeProjectIdx = projects.length - 1;
    }
    renderTopBar();
    
    // Show Undo Button
    const undoBtn = document.getElementById('btnRestoreSet');
    if(undoBtn) {
        undoBtn.style.display = 'inline-flex';
        undoBtn.innerHTML = `<i class="fas fa-undo"></i> Restore "${deletedItem.name}"`;
    }

    if(isOverviewMode) showOverview();
    else switchProject(activeProjectIdx);
}

function restoreProject() {
    if (!deletedItem) return;
    
    // Insert back at original position (or end if out of bounds)
    if (deletedItemIdx >= 0 && deletedItemIdx <= projects.length) {
        projects.splice(deletedItemIdx, 0, deletedItem);
    } else {
        projects.push(deletedItem);
    }
    
    deletedItem = null;
    deletedItemIdx = -1;
    
    // Hide Undo Button
    document.getElementById('btnRestoreSet').style.display = 'none';
    
    renderTopBar();
    // Switch to restored project
    if(isOverviewMode) showOverview();
    else switchProject(projects.indexOf(deletedItem) > -1 ? projects.indexOf(deletedItem) : projects.length - 1);
}

// --- DYNAMIC MODAL ---
function showModal(title, content, type = 'success') {
    const modal = document.getElementById('customModal');
    const titleEl = document.getElementById('modalTitle');
    const msgEl = document.getElementById('modalMsg');
    const iconEl = document.getElementById('modalIcon');
    const btnEl = document.getElementById('modalBtn');
    
    if (modal && titleEl && msgEl) {
        titleEl.innerText = title;
        msgEl.innerHTML = content;
        iconEl.className = 'fas sa-icon-check'; 
        
        if (type === 'error') {
            iconEl.classList.remove('fa-check-circle');
            iconEl.classList.add('fa-times-circle');
            iconEl.style.color = '#ef4444'; 
            btnEl.style.backgroundColor = '#ef4444';
            btnEl.innerText = 'Close';
        } else {
            iconEl.classList.add('fa-check-circle');
            iconEl.classList.remove('fa-times-circle');
            iconEl.style.color = '#10b981';
            btnEl.style.backgroundColor = '#2563eb';
            btnEl.innerText = 'OK';
        }
        modal.classList.add('open');
    } else {
        alert(title + "\n" + content.replace(/<br>/g,'\n').replace(/<[^>]*>?/gm, ''));
    }
}

function closeModal() {
    document.getElementById('customModal')?.classList.remove('open');
}
window.onclick = function(event) {
    if (event.target === document.getElementById('customModal')) closeModal();
}

// --- TOP BAR & OVERVIEW ---
function renderTopBar() {
    const list = document.getElementById('projectTabs');
    list.innerHTML = "";
    projects.forEach((p, i) => {
        const activeClass = (i === activeProjectIdx && !isOverviewMode) ? 'active' : '';
        const statusClass = p.status; 
        list.innerHTML += `
            <div class="tab-item ${activeClass}" onclick="switchProject(${i})">
                <div class="tab-dot ${statusClass}"></div>
                <span>${p.name}</span>
                <button class="btn-tab-close" onclick="deleteProject(event, ${i})">×</button>
            </div>
        `;
    });
    updateAddButtonText();
}

function showOverview() {
    saveCurrentViewToProject();
    isOverviewMode = true;
    
    document.querySelector('.sidebar').classList.add('overview-locked');
    document.getElementById('tabOverview').classList.add('active');
    
    document.querySelectorAll('.tab-item').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById('overviewSection').style.display = 'block';
    
    let totalSets = projects.length, totalRows = 0, totalMatches = 0, totalMismatches = 0, tbody = "";

    projects.forEach(p => {
        const rows = (p.dataA && p.dataB) ? Math.max(p.dataA.body.length, p.dataB.body.length) : 0;
        const matches = p.summary ? p.summary.matches : 0;
        const mismatches = p.summary ? p.summary.mismatches : 0;

        totalRows += rows; 
        totalMatches += matches; 
        totalMismatches += mismatches;
        
        tbody += `<tr>
                <td><strong>${p.name}</strong></td>
                <td><span class="status-dot ${p.status}"></span> ${p.status.toUpperCase()}</td>
                <td>${p.dataA ? p.dataA.body.length : 0}</td>
                <td>${p.dataB ? p.dataB.body.length : 0}</td>
                <td style="color:#10b981; font-weight:bold;">${matches}</td>
                <td style="color:#ef4444; font-weight:bold;">${mismatches}</td>
                <td><button onclick="switchProject(${projects.indexOf(p)})" class="btn-ghost" style="font-size:12px;">View</button></td>
            </tr>`;
    });

    document.getElementById('overviewBody').innerHTML = tbody;
    document.getElementById('overviewStats').innerHTML = `
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${totalSets}</div><div class="bs-lbl">Sets</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${totalRows}</div><div class="bs-lbl">Total Rows</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#10b981">${totalMatches}</div><div class="bs-lbl">Matches</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#ef4444">${totalMismatches}</div><div class="bs-lbl">Mismatches</div></div>
    `;
}

function switchProject(idx) {
    if(isOverviewMode) {
        document.getElementById('tabOverview').classList.remove('active');
        document.querySelector('.sidebar').classList.remove('overview-locked');
        isOverviewMode = false;
    } else { saveCurrentViewToProject(); }
    
    activeProjectIdx = idx;
    
    // Clear Row Undo History when switching projects (to avoid confusion)
    deletedRowsHistory = [];
    document.getElementById('btnUndoRow').style.display = 'none';
    
    loadProjectIntoView(idx);
    renderTopBar();
}

function saveCurrentViewToProject() {
    if (projects.length === 0 || isOverviewMode) return;
    const p = projects[activeProjectIdx];
    if (document.getElementById('step1').style.display !== 'none') {
        p.rawA = document.getElementById('tableA').value;
        p.rawB = document.getElementById('tableB').value;
        p.rawMatrix = document.getElementById('matrixRawInput').value;
        p.matrix = getMatrixDataFromUI(); 
        const matSec = document.getElementById('matrixSection');
        p.showMatrix = (matSec && matSec.style.display !== 'none');
    }
}

function getMatrixDataFromUI() {
    const rows = document.querySelectorAll('.matrix-row-item'); 
    const data = [];
    rows.forEach(r => {
        const k = r.querySelector('.m-key').value.trim();
        const v = r.querySelector('.m-val').value.trim();
        if(k) data.push({key: k, val: v});
    });
    return data;
}

function loadProjectIntoView(idx) {
    const p = projects[idx];
    document.getElementById('tableA').value = p.rawA || "";
    document.getElementById('tableB').value = p.rawB || "";
    document.getElementById('matrixRawInput').value = p.rawMatrix || "";
    const matrixList = document.getElementById('matrixList');
    matrixList.innerHTML = "";
    if (p.matrix && p.matrix.length > 0) p.matrix.forEach(m => addMatrixRow(m.key, m.val));
    else addMatrixRow(); 
    
    const matSec = document.getElementById('matrixSection');
    const btn = document.getElementById('btnToggleMatrix');
    if (p.showMatrix) {
        matSec.style.display = 'block';
        btn.innerHTML = `<i class="fas fa-minus-circle"></i> Hide Matrix Rules`;
    } else {
        matSec.style.display = 'none';
        btn.innerHTML = `<i class="fas fa-plus-circle"></i> Show Matrix Rules (Optional)`;
    }
    
    if (p.step === 3) {
        renderMappingTable(); 
    }
    jumpToStep(p.step || 1);
    
    if (p.step === 4 && p.mapping.length > 0) renderDashboard();
    else if (p.step === 2 && p.dataA) renderPreviewTables();
}

function clearCurrentProject() {
    document.getElementById('tableA').value = "";
    document.getElementById('tableB').value = "";
    document.getElementById('matrixRawInput').value = "";
    document.getElementById('matrixList').innerHTML = "";
    addMatrixRow();
    const p = projects[activeProjectIdx];
    p.rawA = ""; p.rawB = ""; p.rawMatrix = "";
    p.dataA = null; p.dataB = null; p.matrix = []; p.mapping = [];
    p.status = 'empty'; p.step = 1; p.showMatrix = false;
    p.summary = { matches: 0, mismatches: 0 };
    document.getElementById('matrixSection').style.display = 'none';
    renderTopBar();
}

function jumpToStep(step) {
    const p = projects[activeProjectIdx];
    p.step = step;
    
    if(step === 3) {
        renderMappingTable();
    }

    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById(`step${step}`).style.display = 'block';
    document.querySelectorAll('.v-step').forEach(el => el.classList.remove('active'));
    document.getElementById(`navStep${step}`).classList.add('active');
}

// --- STEP 1 -> 2: PREVIEW ---
function goToPreview() {
    saveCurrentViewToProject();
    const p = projects[activeProjectIdx];
    if (!p.rawA.trim() || !p.rawB.trim()) {
        showModal("Missing Data", "Please paste data into both Form Data and Sainpase Data tables.", 'error');
        return;
    }
    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";
    p.dataA = parseExcelData(p.rawA, modeA);
    p.dataB = parseExcelData(p.rawB, modeB);
    if (!p.dataA || !p.dataB) {
        showModal("Error Parsing", "Could not parse data. Please check your inputs.", 'error');
        return;
    }
    autoCleanData(p.dataA);
    autoCleanData(p.dataB);
    p.status = 'ready'; p.step = 2;
    renderTopBar(); renderPreviewTables(); jumpToStep(2);
}

function renderPreviewTables() {
    const p = projects[activeProjectIdx];
    if(!p.dataA || !p.dataB) return;
    renderSinglePreview('previewTableA', 'countA', p.dataA, 'A');
    renderSinglePreview('previewTableB', 'countB', p.dataB, 'B');
    
    // Update Row Undo Button Visibility
    const undoBtn = document.getElementById('btnUndoRow');
    if(deletedRowsHistory.length > 0) {
        undoBtn.style.display = 'inline-flex';
        undoBtn.innerText = `Undo Delete (${deletedRowsHistory.length})`;
    } else {
        undoBtn.style.display = 'none';
    }
}

function renderSinglePreview(tableId, countId, data, side) {
    const tbl = document.getElementById(tableId);
    document.getElementById(countId).innerText = data.body.length;
    let html = "<thead><tr><th>Action</th><th>#</th>";
    data.headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    data.body.forEach((row, idx) => {
        html += `<tr><td><button class="btn-ghost" style="color:red; font-size:11px;" onclick="deleteRow('${side}', ${idx})"><i class="fas fa-trash"></i></button></td><td>${idx+1}</td>`;
        row.forEach(cell => html += `<td>${cell}</td>`);
        html += `</tr>`;
    });
    html += "</tbody>";
    tbl.innerHTML = html;
}

// --- UNDO ROW LOGIC ---
function deleteRow(side, rowIndex) {
    const p = projects[activeProjectIdx];
    let rowData;
    
    if (side === 'A') {
        rowData = p.dataA.body[rowIndex];
        p.dataA.body.splice(rowIndex, 1);
    } else {
        rowData = p.dataB.body[rowIndex];
        p.dataB.body.splice(rowIndex, 1);
    }
    
    // Add to history
    deletedRowsHistory.push({ side: side, idx: rowIndex, data: rowData });
    
    renderPreviewTables();
}

function restoreLastRow() {
    if (deletedRowsHistory.length === 0) return;
    
    const last = deletedRowsHistory.pop();
    const p = projects[activeProjectIdx];
    
    if (last.side === 'A') {
        p.dataA.body.splice(last.idx, 0, last.data);
    } else {
        p.dataB.body.splice(last.idx, 0, last.data);
    }
    
    renderPreviewTables();
}

// --- STEP 2 -> 3: MAPPING ---
function renderMappingTable() {
    const p = projects[activeProjectIdx];
    if (!p.dataA || !p.dataB) return false;

    const tbody = document.getElementById('mappingBody');
    tbody.innerHTML = "";
    const searchInput = document.getElementById('mapSearchInput');
    if(searchInput) searchInput.value = "";

    const leftRows = [];
    p.dataA.headers.forEach((h, i) => { if (h && h.trim() !== "") leftRows.push({ type: 'source', name: h, val: i }); });
    p.matrix.forEach(m => leftRows.push({ type: 'matrix', name: m.key, val: m.key, display: m.val }));

    const rightOptions = p.dataB.headers.map((h, i) => ({ name: h, index: i }));
    rightOptions.sort((a,b) => a.name.localeCompare(b.name));

    leftRows.forEach(row => {
        let matchVal = "-1";
        const savedSession = p.mapping.find(m => {
            if (row.type === 'source') return m.targetType === 'source' && m.targetVal === row.val;
            if (row.type === 'matrix') return m.targetType === 'matrix' && m.targetVal === row.val;
            return false;
        });

        if (savedSession) {
            matchVal = savedSession.idxB;
        } else {
            const memMatch = rightOptions.find(opt => {
                const saved = getSavedMapping(opt.name); 
                if (!saved) return false;
                const [sType, sVal] = saved.split(':');
                if (row.type === 'source' && sType === 'SRC' && sVal === row.name) return true;
                if (row.type === 'matrix' && sType === 'MAT' && sVal === row.name) return true;
                return false;
            });
            if (memMatch) matchVal = memMatch.index;
            else {
                const match = rightOptions.find(t => t.name.toLowerCase() === row.name.toLowerCase());
                if (match) matchVal = match.index;
            }
        }

        const isChecked = matchVal !== "-1" ? "checked" : "";
        const rowClass = matchVal !== "-1" ? "mapped-row" : "";

        let opts = `<option value="-1">-- Ignore / Select --</option>`;
        rightOptions.forEach(opt => {
            const sel = (opt.index === matchVal) ? "selected" : "";
            opts += `<option value="${opt.index}" ${sel}>${opt.name}</option>`;
        });

        let label = row.name;
        if (row.type === 'matrix') label += ` <small style="color:#666">(${row.display})</small>`;

        const tr = document.createElement('tr');
        tr.className = rowClass;
        tr.setAttribute('data-search', row.name.toLowerCase()); 

        tr.innerHTML = `
            <td><b>${label}</b></td>
            <td style="text-align:center"><i class="fas fa-arrow-right" style="color:#9ca3af"></i></td>
            <td><select class="map-select" data-type="${row.type}" data-val="${row.val}" data-name="${row.name}" onchange="autoTick(this)">${opts}</select></td>
            <td style="text-align:center"><input type="checkbox" class="map-check" ${isChecked} onchange="updateMapStats()"></td>
        `;
        tbody.appendChild(tr);
    });
    updateMapStats();
    return true; 
}

function goToMapping() {
    const p = projects[activeProjectIdx];
    if (!p.dataA || !p.dataB) return;
    
    // Strict check for "Next" button
    if (p.dataA.body.length !== p.dataB.body.length) {
        showModal("Row Mismatch!", `Form Data has ${p.dataA.body.length} rows.\nSainpase Data has ${p.dataB.body.length} rows.\nPlease match the number of rows.`, 'error');
        return; 
    }

    renderMappingTable();
    jumpToStep(3);
}

function autoTick(selectEl) { 
    const checkbox = selectEl.parentElement.nextElementSibling.querySelector('input'); 
    const row = selectEl.closest('tr');
    if(selectEl.value !== "-1") {
        checkbox.checked = true;
        row.classList.add('mapped-row');
    } else {
        checkbox.checked = false;
        row.classList.remove('mapped-row');
    }
    updateMapStats(); 
}

function resetMapping() {
    const selects = document.querySelectorAll('.map-select');
    selects.forEach(s => { s.value = "-1"; autoTick(s); });
}

function autoMapAgain() { renderMappingTable(); }

function filterMappingRows() {
    const filter = document.getElementById('mapSearchInput').value.toLowerCase();
    const rows = document.querySelectorAll('#mappingBody tr');
    rows.forEach(r => {
        const txt = r.getAttribute('data-search') || "";
        r.style.display = txt.includes(filter) ? "" : "none";
    });
}

function cleanAllSets() {
    let processed = 0;
    saveCurrentViewToProject(); 
    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";
    projects.forEach(p => {
        if (p.rawA && p.rawB) {
            p.dataA = parseExcelData(p.rawA, modeA);
            p.dataB = parseExcelData(p.rawB, modeB);
            if (p.dataA && p.dataB) {
                autoCleanData(p.dataA);
                autoCleanData(p.dataB);
                p.status = 'ready'; p.step = 2; processed++;
            }
        }
    });
    if (processed === 0) showModal("No Data Found", "No sets with data found to clean.", 'error');
    else {
        showModal("Cleanup Complete", `Cleaned ${processed} sets successfully.`, 'success');
        renderTopBar(); loadProjectIntoView(activeProjectIdx);
    }
}

// --- UPDATED: KEY:VALUE OR KEY=VALUE SUPPORT ---
function autoParseMatrix() {
    const raw = document.getElementById("matrixRawInput").value;
    if (!raw.trim()) return;
    const rawLines = raw.split(/\r?\n/);
    const rows = [];
    let buffer = "";
    rawLines.forEach(line => {
        const l = line.trim();
        if (!l) return;
        buffer += (buffer ? " " : "") + l;
        if (!l.startsWith('"') || l.endsWith('"') || (buffer.startsWith('"') && buffer.includes('"', 1))) {
            rows.push(buffer); buffer = "";
        }
    });
    document.getElementById("matrixList").innerHTML = "";
    rows.forEach(row => {
        let k, v;
        // Supports Tab, Colon (:), or Equals (=)
        if (row.includes("\t")) {
            [k, v] = row.split("\t");
        } else if (row.includes(":")) {
            let idx = row.indexOf(":");
            k = row.substring(0, idx);
            v = row.substring(idx+1);
        } else if (row.includes("=")) { // Added support for =
            let idx = row.indexOf("=");
            k = row.substring(0, idx);
            v = row.substring(idx+1);
        } else {
            k = row;
            v = "";
        }
        addMatrixRow(mapKey(cleanLabel(k || "")), normalizeValue(v || ""));
    });
}

function cleanLabel(text) { return text.replace(/"/g, "").replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim(); }
function normalizeValue(val) { return val ? val.trim() : ""; }
function mapKey(label) {
    const L = label.toUpperCase();
    if (L.startsWith("ADAPTIVE")) return "ADAPTIVE";
    if (L.startsWith("UPF")) return "UPF";
    if (L.startsWith("LYCRA")) return "LYCRA";
    if (L.startsWith("SUSTAINABILITY")) return "SUSTAINABILITY";
    if (L.startsWith("SKU")) return "SKU#";
    return label;
}

function autoCleanData(data) { 
    if (!data || !data.body) return; 
    const qtyIndices = [];
    data.headers.forEach((h, i) => { if (/^(qty|quantity|total\s*qty|total\s*quantity)$/i.test(h.trim())) qtyIndices.push(i); });
    data.body = data.body.filter(row => { 
        const hasContent = row.some(cell => cell && cell.toString().trim() !== ""); 
        if (!hasContent) return false; 
        if (row.some(c => c.toString().includes("UPC") || c.toString().includes("(optional)"))) return false;
        for (let idx of qtyIndices) if ((row[idx] || "").toString().trim() === "" || (row[idx] || "").toString().trim() === "0") return false; 
        return true; 
    }); 
}

function parseExcelData(raw, mode) { 
    if (!raw.trim()) return null; 
    const rowRegex = /\r?\n(?=(?:[^"]*"[^"]*")*[^"]*$)/; 
    let rows = raw.trim().split(rowRegex); 
    while (rows.length > 0 && rows[rows.length - 1].trim() === "") rows.pop(); 
    if (rows.length < 1) return null; 
    let ignoreFirstColumn = false; 
    if (rows.length >= 2 && rows[0].toLowerCase().includes("sku information")) { rows.shift(); ignoreFirstColumn = true; } 
    const delimiter = rows[0].includes('\t') ? '\t' : ','; 
    const splitRow = (r) => r.split(delimiter).map(c => { 
        let val = c.trim(); 
        if (val.startsWith('"') && val.endsWith('"')) val = val.slice(1, -1).replace(/""/g, '"');
        return val; 
    }); 
    let headers = [], bodyStartIndex = 1; 
    if (mode === "2rows") { 
        if (rows.length < 2) return null; 
        const row1 = splitRow(rows[0]), row2 = splitRow(rows[1]); 
        const maxCols = Math.max(row1.length, row2.length); 
        for (let i = 0; i < maxCols; i++) headers.push(cleanHeader(row2[i] || "") || cleanHeader(row1[i] || "")); 
        bodyStartIndex = 2; 
    } else { 
        headers = splitRow(rows[0]).map(h => cleanHeader(h)); bodyStartIndex = 1; 
    } 
    let bodyRaw = rows.slice(bodyStartIndex), body = bodyRaw.map(r => splitRow(r)); 
    if (ignoreFirstColumn) { headers.shift(); body = body.map(r => r.slice(1)); } 
    return { headers, body }; 
}

function cleanHeader(text) { 
    if (!text) return ""; 
    let clean = text.trim(); 
    if (clean.startsWith('"') && clean.endsWith('"')) clean = clean.slice(1, -1); 
    clean = clean.replace(/[\r\n]+/g, " ").replace(/\s*\([^\)]*\)/g, "").replace(/\*/g, "").trim(); 
    if(clean.includes("SIZEStyle")) return "Style"; 
    return clean; 
}

function addMatrixRow(key = "", val = "") { 
    const div = document.createElement('div'); div.className = 'matrix-row-item'; 
    div.innerHTML = `<input type="text" class="matrix-input m-key" placeholder="Field" value="${key}"><input type="text" class="matrix-input m-val" placeholder="Value" value="${val}"><button class="btn-x" onclick="this.parentElement.remove()">×</button>`; 
    document.getElementById('matrixList').appendChild(div); 
}

function toggleMatrix() { 
    const sec = document.getElementById('matrixSection'), btn = document.getElementById('btnToggleMatrix'); 
    if (sec.style.display === 'none') { sec.style.display = 'block'; btn.innerHTML = `<i class="fas fa-minus-circle"></i> Hide Matrix Rules`; } 
    else { sec.style.display = 'none'; btn.innerHTML = `<i class="fas fa-plus-circle"></i> Show Matrix Rules (Optional)`; } 
}

function updateMapStats() { 
    const p = projects[activeProjectIdx]; if (!p.dataB) return; 
    const selectedIndices = new Set(); 
    document.querySelectorAll('.map-select').forEach((sel, i) => { if (document.querySelectorAll('.map-check')[i].checked && sel.value !== "-1") selectedIndices.add(parseInt(sel.value)); }); 
    const unmappedNames = []; 
    p.dataB.headers.forEach((h, i) => { if (!selectedIndices.has(i)) unmappedNames.push(h); }); 
    const display = document.getElementById('unmappedList'); 
    if (display) { 
        if (unmappedNames.length === 0) { display.innerText = "All Mapped ✓"; display.style.background = "#d1fae5"; display.style.color = "#065f46"; } 
        else { display.innerText = unmappedNames.join(", "); display.style.background = "#fee2e2"; display.style.color = "#b91c1c"; } 
    } 
}

function generateResults() { 
    saveMappingFromUI(); 
    const p = projects[activeProjectIdx]; 
    
    // Strict Row Check
    if (p.dataA.body.length !== p.dataB.body.length) {
         showModal("Row Mismatch", "Cannot run analysis.\nRow counts must be equal.", 'error');
         return;
    }

    if (!p.mapping || p.mapping.length === 0) { showModal("Error", "Please map at least one column.", 'error'); return; } 
    
    p.summary = calculateStats(p);

    p.status = 'done'; p.step = 4; 
    renderTopBar(); 
    renderDashboard(); 
    jumpToStep(4); 
    
    // Show Mismatch Report Popup
    showAnalysisReport(p);
}

function saveMappingFromUI() { 
    const p = projects[activeProjectIdx]; 
    const selects = document.querySelectorAll('.map-select'); 
    const checks = document.querySelectorAll('.map-check'); 
    p.mapping = []; 
    selects.forEach((sel, i) => { 
        if (checks[i].checked && sel.value !== "-1") { 
            const targetIdx = parseInt(sel.value); 
            const srcType = sel.getAttribute('data-type'), srcVal = sel.getAttribute('data-val'), srcName = sel.getAttribute('data-name'); 
            p.mapping.push({ name: p.dataB.headers[targetIdx], idxB: targetIdx, targetType: srcType, targetVal: (srcType === 'source' ? parseInt(srcVal) : srcVal) }); 
            saveMappingMemory(p.dataB.headers[targetIdx], srcType, srcName); 
        } 
    }); 
}

// --- UPDATED SMART RUN LOGIC ---
function runAllComparisons() { 
    saveCurrentViewToProject(); 
    let processed = 0;
    let skipped = 0;
    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";
    
    projects.forEach(p => {
        if (p.rawA && p.rawB) {
            if (!p.dataA || !p.dataB) { 
                p.dataA = parseExcelData(p.rawA, modeA); 
                p.dataB = parseExcelData(p.rawB, modeB); 
                autoCleanData(p.dataA); 
                autoCleanData(p.dataB); 
            }
            if (p.dataA && p.dataB) { 
                // Check Row Mismatch
                if (p.dataA.body.length === p.dataB.body.length) {
                    autoMapProject(p);
                    p.summary = calculateStats(p);
                    p.status = 'done'; 
                    p.step = 4; 
                    processed++;
                } else {
                    skipped++; 
                }
            }
        }
    });

    if (processed === 0 && skipped === 0) {
        showModal("No Data", "No sets with data found.", 'error'); 
        return;
    }

    // 1. Single Set Logic
    if (projects.length === 1) {
        if (skipped > 0) {
            const p = projects[0];
            showModal("Row Mismatch", `Analysis Failed.\n\nForm Data: ${p.dataA.body.length} rows\nSainpase Data: ${p.dataB.body.length} rows\n\nPlease fix the mismatch in this set.`, 'error');
            return; 
        }
        
        if (processed === 1) {
            activeProjectIdx = 0;
            loadProjectIntoView(0); 
            jumpToStep(4);          
            renderTopBar();
            showAnalysisReport(projects[0]);
            return; 
        }
    }

    // 2. Multi-Set Logic
    if (skipped > 0) {
        showModal("Attention Needed", `Analysis ran for ${processed} set(s).\n\nSkipped ${skipped} set(s) due to row mismatches.\nCheck the Overview for details.`, 'error'); 
    }

    renderTopBar(); 
    showOverview(); 
}

function calculateStats(p) {
    if (!p.dataA || !p.dataB || !p.mapping || p.mapping.length === 0) return { matches: 0, mismatches: 0 };
    
    let matches = 0, mismatches = 0;
    const rows = Math.max(p.dataA.body.length, p.dataB.body.length);
    
    p.mapping.forEach(map => {
        const lowerName = map.name.toLowerCase();
        const isPrice = /price|cost|retail/i.test(lowerName);
        const isQty = /qty|quantity/i.test(lowerName);
        for(let i=0; i<rows; i++) {
            let vB = (p.dataB.body[i]?.[map.idxB] || "").trim();
            let vA = map.targetType === 'matrix' 
                ? (p.matrix.find(m => m.key === map.targetVal)?.val || "")
                : (p.dataA.body[i]?.[map.targetVal] || "").trim();
            
            // Exact Match Logic
            let normA = String(vA).trim().toLowerCase();
            let normB = String(vB).trim().toLowerCase();
            
            let equal = false;
            
            if (isPrice) equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,''));
            else if (isQty) equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,''));
            else equal = (normA === normB);
            
            if(equal) matches++; else mismatches++;
        }
    });
    return { matches, mismatches };
}

function showAnalysisReport(p) {
    const stats = p.summary || calculateStats(p);
    
    if (stats.mismatches === 0) {
        showModal("Analysis Complete: Perfect Match!", 
            `<div style="text-align:center;">
                <p style="font-size:16px; margin-bottom:10px;">All <strong>${p.dataA.body.length}</strong> rows match perfectly across all columns.</p>
                <div style="background:#f0fdf4; color:#166534; padding:15px; border-radius:8px; font-weight:bold;">
                    NO MISMATCHES FOUND
                </div>
            </div>`, 
            'success'
        );
    } else {
        let listHtml = `<table style="width:100%; border-collapse:collapse; margin-top:15px; font-size:13px;">
            <thead>
                <tr style="background:#fef2f2; color:#b91c1c; border-bottom:1px solid #ef4444;">
                    <th style="padding:8px; text-align:left;">Column Name</th>
                    <th style="padding:8px; text-align:right;">Mismatches</th>
                </tr>
            </thead>
            <tbody>`;
        
        p.mapping.forEach(map => {
            let missCount = 0;
            const rows = p.dataA.body.length;
            const lowerName = map.name.toLowerCase();
            const isPrice = /price|cost|retail/i.test(lowerName);
            const isQty = /qty|quantity/i.test(lowerName);

            for(let i=0; i<rows; i++) {
                 let vB = (p.dataB.body[i]?.[map.idxB] || "").trim();
                 let vA = map.targetType === 'matrix' 
                    ? (p.matrix.find(m => m.key === map.targetVal)?.val || "")
                    : (p.dataA.body[i]?.[map.targetVal] || "").trim();

                 let normA = String(vA).trim().toLowerCase();
                 let normB = String(vB).trim().toLowerCase();

                 let equal = false;
                 if (isPrice) equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,''));
                 else if (isQty) equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,''));
                 else equal = (normA === normB);
                 
                 if(!equal) missCount++;
            }

            if(missCount > 0) {
                listHtml += `<tr>
                    <td style="padding:8px; border-bottom:1px solid #fee2e2;">${map.name}</td>
                    <td style="padding:8px; border-bottom:1px solid #fee2e2; text-align:right; font-weight:bold; color:#ef4444;">${missCount}</td>
                </tr>`;
            }
        });
        listHtml += `</tbody></table>`;

        showModal("Analysis Complete: Mismatches Found", 
            `<div style="text-align:left;">
                <p>Found <strong>${stats.mismatches}</strong> total mismatches in <strong>${p.dataA.body.length}</strong> rows.</p>
                ${listHtml}
            </div>`, 
            'error'
        );
    }
}

function autoMapProject(p) { 
    p.mapping = []; 
    p.matrix.forEach(m => { 
        let targetIdx = -1; const memTarget = p.dataB.headers.find(h => getSavedMapping(h) === `MAT:${m.key}`); 
        if (memTarget) targetIdx = p.dataB.headers.indexOf(memTarget); else targetIdx = p.dataB.headers.findIndex(h => h.toLowerCase() === m.key.toLowerCase()); 
        if (targetIdx !== -1) p.mapping.push({ name: p.dataB.headers[targetIdx], idxB: targetIdx, targetType: 'matrix', targetVal: m.key }); 
    }); 
    p.dataA.headers.forEach((hA, iA) => { 
        let targetIdx = -1; const memTarget = p.dataB.headers.find(h => getSavedMapping(h) === `SRC:${hA}`); 
        if (memTarget) targetIdx = p.dataB.headers.indexOf(memTarget); else targetIdx = p.dataB.headers.findIndex(h => h.toLowerCase() === hA.toLowerCase()); 
        if (targetIdx !== -1) p.mapping.push({ name: p.dataB.headers[targetIdx], idxB: targetIdx, targetType: 'source', targetVal: iA }); 
    }); 
}

function renderDashboard() { 
    const p = projects[activeProjectIdx]; 
    const stats = p.summary || calculateStats(p);
    
    const maxRows = Math.max(p.dataA.body.length, p.dataB.body.length); 
    const cardArea = document.getElementById('summaryCards'); cardArea.innerHTML = ""; 
    
    p.mapping.forEach(map => { 
        let match = 0, miss = 0; 
        const lowerName = map.name.toLowerCase(); 
        const isPrice = /price|cost|retail/i.test(lowerName), isQty = /qty|quantity/i.test(lowerName); 
        for (let i = 0; i < maxRows; i++) { 
            let vB = (p.dataB.body[i]?.[map.idxB] || "").trim(); 
            let vA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").trim(); 
            let normA = String(vA).trim().toLowerCase(), normB = String(vB).trim().toLowerCase(); 
            let equal = false; 
            if (isPrice) equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,'')); 
            else if (isQty) equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,'')); 
            else equal = (normA === normB); 
            if (equal) match++; else miss++; 
        } 
        
        const cls = miss > 0 ? 'bg-warn' : 'bg-good'; 
        cardArea.innerHTML += `<div class="field-card ${cls}"><div class="fc-head">${map.name}</div><div class="fc-stats"><span style="color:#10b981">✓ ${match}</span><span style="color:#ef4444">✗ ${miss}</span></div></div>`; 
    }); 
    
    document.getElementById('globalStats').innerHTML = `<div class="big-stat"><div class="bs-val" style="color:#2563eb">${maxRows}</div><div class="bs-lbl">Rows</div></div><div class="big-stat"><div class="bs-val" style="color:#2563eb">${p.mapping.length}</div><div class="bs-lbl">Fields</div></div><div class="big-stat"><div class="bs-val" style="color:#10b981">${stats.matches}</div><div class="bs-lbl">Matches</div></div><div class="big-stat"><div class="bs-val" style="color:#ef4444">${stats.mismatches}</div><div class="bs-lbl">Mismatches</div></div>`; 
    renderResultTables(maxRows); 
}

function renderResultTables(maxRows) { 
    const p = projects[activeProjectIdx]; 
    const tA = document.getElementById('renderTableA'), tB = document.getElementById('renderTableB'); 
    let hA = "<thead><tr><th>#</th>" + p.dataA.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    let hB = "<thead><tr><th>#</th>" + p.dataB.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    let bA = "", bB = ""; 
    const mapLookup = {}; p.mapping.forEach(m => mapLookup[m.idxB] = m); 
    const reverseLookup = {}; p.mapping.forEach(m => { if (m.targetType === 'source') { if (!reverseLookup[m.targetVal]) reverseLookup[m.targetVal] = []; reverseLookup[m.targetVal].push(m); } }); 
    for (let i = 0; i < maxRows; i++) { 
        let rA = `<td>${i+1}</td>`; 
        for (let cA = 0; cA < p.dataA.headers.length; cA++) { 
            let vA = (p.dataA.body[i]?.[cA] || "").trim(), cls = ""; 
            if (reverseLookup[cA]) { 
                let allMatch = true; 
                reverseLookup[cA].forEach(map => { 
                    let vB = (p.dataB.body[i]?.[map.idxB] || "").trim(); 
                    let normA = String(vA).trim().toLowerCase(), normB = String(vB).trim().toLowerCase(); 
                    const lowerName = map.name.toLowerCase(), isPrice = /price|cost|retail/i.test(lowerName), isQty = /qty|quantity/i.test(lowerName); 
                    let equal = false; 
                    if (isPrice) equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,'')); 
                    else if (isQty) equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,'')); 
                    else equal = (normA === normB); 
                    if (!equal) allMatch = false; 
                }); 
                cls = allMatch ? "match" : "diff"; 
            } 
            rA += `<td class="${cls}">${vA}</td>`; 
        } 
        let rB = `<td>${i+1}</td>`; 
        p.dataB.headers.forEach((_, colIdx) => { 
            let vB = (p.dataB.body[i]?.[colIdx] || "").trim(), cls = ""; 
            if (mapLookup.hasOwnProperty(colIdx)) { 
                let map = mapLookup[colIdx], vA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").trim(); 
                let normA = String(vA).trim().toLowerCase(), normB = String(vB).trim().toLowerCase(); 
                const lowerName = p.dataB.headers[colIdx].toLowerCase(), isPrice = /price|cost|retail/i.test(lowerName), isQty = /qty|quantity/i.test(lowerName); 
                let equal = false; 
                if (isPrice) equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,'')); 
                else if (isQty) equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,'')); 
                else equal = (normA === normB); 
                cls = equal ? "match" : "diff"; 
            } 
            rB += `<td class="${cls}">${vB}</td>`; 
        }); 
        bA += `<tr>${rA}</tr>`; bB += `<tr id="rowB-${i}">${rB}</tr>`; 
    } 
    tA.innerHTML = hA + bA + "</tbody>"; tB.innerHTML = hB + bB + "</tbody>"; 
}

function getSavedMapping(headerName) { return localStorage.getItem("map_" + headerName); }
function saveMappingMemory(headerName, type, val) { localStorage.setItem("map_" + headerName, `${type === 'source' ? 'SRC' : 'MAT'}:${val}`); }