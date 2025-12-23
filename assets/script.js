// --- GLOBAL VARIABLES ---
let projects = [];
let activeProjectIdx = 0;
let isOverviewMode = false;
const MAX_TOTAL_SETS = 50; 
let deletedItem = null;
let deletedItemIdx = -1;
let toastTimeout;

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

// --- PROJECT MANAGEMENT (Add, Delete, Switch) ---

function addSetsFromInput() {
    const input = document.getElementById('addSetQty');
    let qty = parseInt(input.value) || 1;
    if (qty < 1) qty = 1;
    if (projects.length + qty > MAX_TOTAL_SETS) {
        alert(`Cannot add. Limit is ${MAX_TOTAL_SETS} sets.`);
        return;
    }
    for(let i=0; i<qty; i++) {
        createSet();
    }
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
        mapping: [], step: 1,
        showMatrix: false 
    });
}

function deleteProject(e, idx) {
    e.stopPropagation(); 
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
    if(isOverviewMode) showOverview();
    else switchProject(activeProjectIdx);
    
    showToast();
}

function showToast() {
    const t = document.getElementById('toast');
    t.classList.remove('hidden');
    clearTimeout(toastTimeout);
    toastTimeout = setTimeout(() => { t.classList.add('hidden'); deletedItem = null; }, 5000); 
}

function undoDelete() {
    if (!deletedItem) return;
    if (deletedItemIdx >= 0 && deletedItemIdx <= projects.length) {
         projects.splice(deletedItemIdx, 0, deletedItem);
    } else {
         projects.push(deletedItem);
    }
    renderTopBar();
    switchProject(deletedItemIdx !== -1 ? deletedItemIdx : projects.length - 1);
    document.getElementById('toast').classList.add('hidden');
    deletedItem = null;
}

// --- TOP BAR & OVERVIEW RENDERER ---

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
    
    document.getElementById('tabOverview').classList.add('active');
    document.querySelectorAll('.tab-item').forEach(el => el.classList.remove('active'));
    
    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById('overviewSection').style.display = 'block';
    
    let totalSets = projects.length;
    let totalRows = 0;
    let totalMatches = 0;
    let totalMismatches = 0;
    let tbody = "";

    projects.forEach(p => {
        let matches = 0, mismatches = 0;
        let rows = 0;
        
        if (p.step === 4 && p.mapping.length > 0) {
            rows = Math.max(p.dataA.body.length, p.dataB.body.length);
            p.mapping.forEach(map => {
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
                    
                    if(equal) matches++; else mismatches++;
                }
            });
        }

        totalRows += rows;
        totalMatches += matches;
        totalMismatches += mismatches;

        tbody += `
            <tr>
                <td><strong>${p.name}</strong></td>
                <td><span class="status-dot ${p.status}" style="display:inline-block"></span> ${p.status.toUpperCase()}</td>
                <td>${p.dataA ? p.dataA.body.length : 0}</td>
                <td>${p.dataB ? p.dataB.body.length : 0}</td>
                <td style="color:#059669">${matches}</td>
                <td style="color:#dc2626">${mismatches}</td>
                <td><button onclick="switchProject(${projects.indexOf(p)})" class="btn-ghost" style="font-size:11px;">View</button></td>
            </tr>
        `;
    });

    document.getElementById('overviewBody').innerHTML = tbody;
    document.getElementById('overviewStats').innerHTML = `
        <div class="big-stat blue"><div class="bs-val">${totalSets}</div><div class="bs-lbl">Sets</div></div>
        <div class="big-stat blue"><div class="bs-val">${totalRows}</div><div class="bs-lbl">Total Rows</div></div>
        <div class="big-stat green"><div class="bs-val">${totalMatches}</div><div class="bs-lbl">Matches</div></div>
        <div class="big-stat red"><div class="bs-val">${totalMismatches}</div><div class="bs-lbl">Mismatches</div></div>
    `;
}

function switchProject(idx) {
    if(isOverviewMode) {
        document.getElementById('tabOverview').classList.remove('active');
        isOverviewMode = false;
    } else {
        saveCurrentViewToProject();
    }
    
    activeProjectIdx = idx;
    loadProjectIntoView(idx);
    renderTopBar();
}

// --- STATE MANAGEMENT (Save/Load) ---

function saveCurrentViewToProject() {
    if (projects.length === 0 || isOverviewMode) return;
    const p = projects[activeProjectIdx];
    // Only save inputs if we are on Step 1
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
    const rows = document.querySelectorAll('.matrix-row');
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
    
    // Load Step 1 Inputs
    document.getElementById('tableA').value = p.rawA || "";
    document.getElementById('tableB').value = p.rawB || "";
    document.getElementById('matrixRawInput').value = p.rawMatrix || "";
    
    // Matrix UI
    const matrixList = document.getElementById('matrixList');
    matrixList.innerHTML = "";
    if (p.matrix && p.matrix.length > 0) {
        p.matrix.forEach(m => addMatrixRow(m.key, m.val));
    } else {
        addMatrixRow(); 
    }
    
    const matSec = document.getElementById('matrixSection');
    const btn = document.getElementById('btnToggleMatrix');
    if (p.showMatrix) {
        matSec.style.display = 'block';
        btn.innerText = "- Hide Matrix Rules";
    } else {
        matSec.style.display = 'none';
        btn.innerText = "+ Show Matrix Rules (Optional)";
    }

    jumpToStep(p.step || 1);
    
    // If specific steps are ready, trigger their renderers
    if (p.step === 4 && p.mapping.length > 0) {
        renderDashboard();
    } else if (p.step === 2 && p.dataA) {
        renderPreviewTables();
    } else if (p.step === 3) {
        goToMapping(); 
    }
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
    p.status = 'empty';
    p.step = 1;
    p.showMatrix = false;
    
    document.getElementById('matrixSection').style.display = 'none';
    document.getElementById('btnToggleMatrix').innerText = "+ Show Matrix Rules (Optional)";
    renderTopBar();
}

function jumpToStep(step) {
    const p = projects[activeProjectIdx];
    p.step = step;
    
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
        alert("Please paste data into both Source and Extracted tables.");
        return;
    }

    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";

    // Parse Data
    p.dataA = parseExcelData(p.rawA, modeA);
    p.dataB = parseExcelData(p.rawB, modeB);

    if (!p.dataA || !p.dataB) {
        alert("Could not parse data. Please check inputs.");
        return;
    }

    // Auto Clean 
    autoCleanData(p.dataA);
    autoCleanData(p.dataB);

    p.status = 'ready';
    p.step = 2;
    renderTopBar();
    renderPreviewTables();
    jumpToStep(2);
}

// UPDATE: Delete Button moved to start (left)
function renderPreviewTables() {
    const p = projects[activeProjectIdx];
    if(!p.dataA || !p.dataB) return;

    renderSinglePreview('previewTableA', 'countA', p.dataA, 'A');
    renderSinglePreview('previewTableB', 'countB', p.dataB, 'B');
}

function renderSinglePreview(tableId, countId, data, side) {
    const tbl = document.getElementById(tableId);
    const countSpan = document.getElementById(countId);
    
    countSpan.innerText = data.body.length;
    
    // Header: Added 'Action' at the START
    let html = "<thead><tr><th>Action</th><th>#</th>";
    data.headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    
    data.body.forEach((row, idx) => {
        html += `<tr>`;
        // Body: Added Delete Button at the START
        html += `<td><button class="btn-del" onclick="deleteRow('${side}', ${idx})">Delete</button></td>`;
        html += `<td>${idx+1}</td>`;
        row.forEach(cell => html += `<td>${cell}</td>`);
        html += `</tr>`;
    });
    html += "</tbody>";
    
    tbl.innerHTML = html;
}

function deleteRow(side, rowIndex) {
    const p = projects[activeProjectIdx];
    if (side === 'A') {
        p.dataA.body.splice(rowIndex, 1);
    } else {
        p.dataB.body.splice(rowIndex, 1);
    }
    renderPreviewTables();
}

// --- STEP 2 -> 3: MAPPING ---

function goToMapping() {
    const p = projects[activeProjectIdx];
    const tbody = document.getElementById('mappingBody');
    tbody.innerHTML = "";
    
    const leftRows = [];
    // Only valid headers
    p.dataA.headers.forEach((h, i) => { if (h && h.trim() !== "") leftRows.push({ type: 'source', name: h, val: i }); });
    p.matrix.forEach(m => leftRows.push({ type: 'matrix', name: m.key, val: m.key, display: m.val }));

    const rightOptions = p.dataB.headers.map((h, i) => ({ name: h, index: i }));
    rightOptions.sort((a,b) => a.name.localeCompare(b.name));

    leftRows.forEach(row => {
        let matchVal = "-1";
        
        // Priority 1: Existing Session Mapping
        const savedSession = p.mapping.find(m => {
            if (row.type === 'source') return m.targetType === 'source' && m.targetVal === row.val;
            if (row.type === 'matrix') return m.targetType === 'matrix' && m.targetVal === row.val;
            return false;
        });

        if (savedSession) {
            matchVal = savedSession.idxB;
        } else {
            // Priority 2: Memory (Local Storage)
            const memMatch = rightOptions.find(opt => {
                const saved = getSavedMapping(opt.name); 
                if (!saved) return false;
                const [sType, sVal] = saved.split(':');
                if (row.type === 'source' && sType === 'SRC' && sVal === row.name) return true;
                if (row.type === 'matrix' && sType === 'MAT' && sVal === row.name) return true;
                return false;
            });

            if (memMatch) {
                matchVal = memMatch.index;
            } else {
                // Priority 3: Name Match
                const match = rightOptions.find(t => t.name.toLowerCase() === row.name.toLowerCase());
                if (match) matchVal = match.index;
            }
        }

        const isChecked = matchVal !== "-1" ? "checked" : "";

        let opts = `<option value="-1">-- Ignore / Select --</option>`;
        rightOptions.forEach(opt => {
            const sel = (opt.index === matchVal) ? "selected" : "";
            opts += `<option value="${opt.index}" ${sel}>${opt.name}</option>`;
        });

        let label = row.name;
        if (row.type === 'matrix') label += ` <small style="color:#666">(${row.display})</small>`;

        const typeAttr = `data-type="${row.type}"`;
        const valAttr = `data-val="${row.val}"`; 
        const nameAttr = `data-name="${row.name}"`;

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><b>${label}</b></td>
            <td style="text-align:center">➔</td>
            <td><select class="map-select" ${typeAttr} ${valAttr} ${nameAttr} onchange="autoTick(this)">${opts}</select></td>
            <td style="text-align:center"><input type="checkbox" class="map-check" ${isChecked} onchange="updateMapStats()"></td>
        `;
        tbody.appendChild(tr);
    });

    jumpToStep(3);
    updateMapStats(); 
}

// --- DATA CLEANING & PARSING LOGIC ---

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
                p.status = 'ready'; 
                p.step = 2; 
                processed++;
            }
        }
    });

    if (processed === 0) alert("No sets with data found to clean.");
    else {
        alert(`Cleaned ${processed} sets.`);
        renderTopBar();
        loadProjectIntoView(activeProjectIdx);
    }
}

// --- NEW MATRIX LOGIC (v3.3) ---

function autoParseMatrix() {
    const raw = document.getElementById("matrixRawInput").value;
    if (!raw.trim()) return;

    // Normalize lines
    const rawLines = raw.split(/\r?\n/);
    const rows = [];

    // Merge multiline quoted labels into one row
    let buffer = "";
    rawLines.forEach(line => {
        const l = line.trim();
        if (!l) return;

        buffer += (buffer ? " " : "") + l;

        // Logic from prompt:
        // If it doesn't start with quote OR it ends with quote
        // OR (if buffer started with quote and now contains a second quote indicating closure)
        if (!l.startsWith('"') || l.endsWith('"') || (buffer.startsWith('"') && buffer.includes('"', 1))) {
            rows.push(buffer);
            buffer = "";
        }
    });

    const list = document.getElementById("matrixList");
    list.innerHTML = "";

    rows.forEach(row => {
        const cols = row.split("\t").map(c => c.trim());

        const rawLabel = cols[0] || "";
        const rawValue = cols[1] || "";

        const label = cleanLabel(rawLabel);
        const key = mapKey(label);
        const value = normalizeValue(rawValue);

        addMatrixRow(key, value);
    });
}

function cleanLabel(text) {
    return text
        .replace(/"/g, "")
        .replace(/\(.*?\)/g, "")
        .replace(/\s+/g, " ")
        .trim();
}

function normalizeValue(val) {
    if (!val) return "";
    const v = val.toUpperCase().trim();
    if (v === "YES" || v === "TRUE") return "YES";
    if (v === "NO" || v === "FALSE") return "NO";
    return val;
}

function mapKey(label) {
    const L = label.toUpperCase();

    if (L.startsWith("ADAPTIVE")) return "ADAPTIVE";
    if (L.startsWith("UPF")) return "UPF";
    if (L.startsWith("LYCRA")) return "LYCRA";
    if (L.startsWith("SUSTAINABILITY")) return "SUSTAINABILITY";
    if (L.startsWith("SKU")) return "SKU#";

    if (L === "OPTIONS") return "Options";
    if (L === "DEPT") return "DEPT";
    if (L === "MAJOR CLASS") return "MAJOR CLASS";
    if (L === "SUB CLASS") return "SUB CLASS";
    if (L === "DESCRIPTION") return "DESCRIPTION";

    return label;
}

// --- END NEW MATRIX LOGIC ---

// UPDATE: Fixed "SIZE" data loss issue & Added Qty check
function autoCleanData(data) { 
    if (!data || !data.body) return; 
    
    // 1. Identify Quantity Columns
    const qtyIndices = [];
    data.headers.forEach((h, i) => {
        // Matches "qty", "quantity", "total qty", "total quantity" (Case Insensitive)
        if (/^(qty|quantity|total\s*qty|total\s*quantity)$/i.test(h.trim())) {
            qtyIndices.push(i);
        }
    });

    // 2. Filter Rows
    data.body = data.body.filter(row => { 
        // A. Delete Empty Rows
        const hasContent = row.some(cell => cell && cell.toString().trim() !== ""); 
        if (!hasContent) return false; 
        
        // B. Delete Garbage Headers (Removed "SIZE" to prevent data loss)
        const isHeaderRepeat = row.some(c => c.toString().includes("UPC") || c.toString().includes("(optional)"));
        if (isHeaderRepeat) return false;

        // C. Delete if Quantity is 0 or Blank
        // Only runs if a 'Quantity' column was actually found
        for (let idx of qtyIndices) {
            const val = (row[idx] || "").toString().trim();
            // Delete if blank or explicitly "0"
            if (val === "" || val === "0") return false; 
        }

        return true; 
    }); 
}

function parseExcelData(raw, mode) { 
    if (!raw.trim()) return null; 
    
    // Robust Regex to split by newline only if NOT inside double quotes
    const rowRegex = /\r?\n(?=(?:[^"]*"[^"]*")*[^"]*$)/; 
    
    let rows = raw.trim().split(rowRegex); 
    
    while (rows.length > 0 && rows[rows.length - 1].trim() === "") rows.pop(); 
    if (rows.length < 1) return null; 

    // Handle "SKU Information" junk header
    let ignoreFirstColumn = false; 
    if (rows.length >= 2 && rows[0].toLowerCase().includes("sku information")) { 
        rows.shift(); 
        ignoreFirstColumn = true; 
    } 

    const delimiter = rows[0].includes('\t') ? '\t' : ','; 
    const splitRow = (r) => r.split(delimiter).map(c => { 
        let val = c.trim(); 
        if (val.startsWith('"') && val.endsWith('"')) {
            val = val.slice(1, -1);
            val = val.replace(/""/g, '"');
        }
        return val; 
    }); 

    let headers = []; 
    let bodyStartIndex = 1; 
    
    if (mode === "2rows") { 
        if (rows.length < 2) return null; 
        const row1 = splitRow(rows[0]); 
        const row2 = splitRow(rows[1]); 
        const maxCols = Math.max(row1.length, row2.length); 
        for (let i = 0; i < maxCols; i++) { 
            const h1 = cleanHeader(row1[i] || ""); 
            const h2 = cleanHeader(row2[i] || ""); 
            headers.push(h2 ? h2 : h1); 
        } 
        bodyStartIndex = 2; 
    } else { 
        const row1 = splitRow(rows[0]); 
        headers = row1.map(h => cleanHeader(h)); 
        bodyStartIndex = 1; 
    } 
    
    let bodyRaw = rows.slice(bodyStartIndex); 
    let body = bodyRaw.map(r => splitRow(r)); 
    
    if (ignoreFirstColumn) { 
        headers.shift(); 
        body = body.map(r => r.slice(1)); 
    } 
    
    return { headers, body }; 
}

function cleanHeader(text) { 
    if (!text) return ""; 
    let clean = text.trim(); 
    if (clean.startsWith('"') && clean.endsWith('"')) clean = clean.slice(1, -1); 
    
    clean = clean.replace(/[\r\n]+/g, " "); 
    
    // Strict BRACKET REMOVAL (Requested)
    clean = clean.replace(/\s*\([^\)]*\)/g, "");

    clean = clean.replace(/\*/g, ""); 
    clean = clean.trim(); 
    
    // Specific fix for "SIZEStyle" issue
    if(clean.includes("SIZEStyle")) return "Style"; 
    
    return clean; 
}

function addMatrixRow(key = "", val = "") { 
    const div = document.createElement('div'); 
    div.className = 'matrix-row'; 
    div.innerHTML = `<input type="text" class="matrix-input m-key" placeholder="Field" value="${key}"><input type="text" class="matrix-input m-val" placeholder="Value" value="${val}"><button class="btn-x" onclick="this.parentElement.remove()">×</button>`; 
    document.getElementById('matrixList').appendChild(div); 
}

function toggleMatrix() { 
    const sec = document.getElementById('matrixSection'); 
    const btn = document.getElementById('btnToggleMatrix'); 
    if (sec.style.display === 'none') { 
        sec.style.display = 'block'; 
        btn.innerText = "- Hide Matrix Rules"; 
    } else { 
        sec.style.display = 'none'; 
        btn.innerText = "+ Show Matrix Rules (Optional)"; 
    } 
}

function autoTick(selectEl) { 
    const checkbox = selectEl.parentElement.nextElementSibling.querySelector('input'); 
    checkbox.checked = (selectEl.value !== "-1"); 
    updateMapStats(); 
}

function updateMapStats() { 
    const p = projects[activeProjectIdx]; 
    if (!p.dataB) return; 
    const selectedIndices = new Set(); 
    document.querySelectorAll('.map-select').forEach((sel, i) => { 
        if (document.querySelectorAll('.map-check')[i].checked && sel.value !== "-1") selectedIndices.add(parseInt(sel.value)); 
    }); 
    const unmappedNames = []; 
    p.dataB.headers.forEach((h, i) => { 
        if (!selectedIndices.has(i)) unmappedNames.push(h); 
    }); 
    const display = document.getElementById('unmappedList'); 
    if (display) { 
        if (unmappedNames.length === 0) { 
            display.innerText = "All Mapped ✓"; 
            display.style.background = "#d1fae5"; 
            display.style.color = "#065f46"; 
        } else { 
            display.innerText = unmappedNames.join(", "); 
            display.style.background = "#fee2e2"; 
            display.style.color = "#b91c1c"; 
        } 
    } 
}

// --- RESULTS & DASHBOARD ---

function generateResults() { 
    saveMappingFromUI(); 
    const p = projects[activeProjectIdx]; 
    if (!p.mapping || p.mapping.length === 0) { 
        alert("Please map at least one column."); 
        return; 
    } 
    p.status = 'done'; 
    p.step = 4; 
    renderTopBar(); 
    renderDashboard(); 
    jumpToStep(4); 
}

function saveMappingFromUI() { 
    const p = projects[activeProjectIdx]; 
    const selects = document.querySelectorAll('.map-select'); 
    const checks = document.querySelectorAll('.map-check'); 
    p.mapping = []; 
    selects.forEach((sel, i) => { 
        if (checks[i].checked && sel.value !== "-1") { 
            const targetIdx = parseInt(sel.value); 
            const srcType = sel.getAttribute('data-type'); 
            const srcVal = sel.getAttribute('data-val'); 
            const srcName = sel.getAttribute('data-name'); 
            p.mapping.push({ 
                name: p.dataB.headers[targetIdx], 
                idxB: targetIdx, 
                targetType: srcType, 
                targetVal: (srcType === 'source' ? parseInt(srcVal) : srcVal) 
            }); 
            saveMappingMemory(p.dataB.headers[targetIdx], srcType, srcName); 
        } 
    }); 
}

// UPDATE: Added saving and redirect to Overview
function runAllComparisons() { 
    // CRITICAL FIX: Save the data currently on the screen before running
    saveCurrentViewToProject(); 

    let processed = 0;
    // Get header modes from UI (defaults to 1row if missing)
    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";

    projects.forEach(p => {
        // Check if raw data exists. If it does, ensure it is parsed.
        if (p.rawA && p.rawB) {
            if (!p.dataA || !p.dataB) {
                p.dataA = parseExcelData(p.rawA, modeA);
                p.dataB = parseExcelData(p.rawB, modeB);
                autoCleanData(p.dataA);
                autoCleanData(p.dataB);
            }

            // Perform the mapping and mark as done
            if (p.dataA && p.dataB) {
                autoMapProject(p); 
                p.status = 'done'; 
                p.step = 4; 
                processed++; 
            }
        }
    });

    if (processed === 0) {
        alert("No sets with data found. Please paste data first.");
    } else {
        renderTopBar();
        showOverview(); // <--- This commands the switch to Overview
    }
}

function autoMapProject(p) { 
    p.mapping = []; 
    p.matrix.forEach(m => { 
        let targetIdx = -1; 
        const memTarget = p.dataB.headers.find(h => getSavedMapping(h) === `MAT:${m.key}`); 
        if (memTarget) targetIdx = p.dataB.headers.indexOf(memTarget); 
        else targetIdx = p.dataB.headers.findIndex(h => h.toLowerCase() === m.key.toLowerCase()); 
        if (targetIdx !== -1) { 
            p.mapping.push({ name: p.dataB.headers[targetIdx], idxB: targetIdx, targetType: 'matrix', targetVal: m.key }); 
        } 
    }); 
    p.dataA.headers.forEach((hA, iA) => { 
        let targetIdx = -1; 
        const memTarget = p.dataB.headers.find(h => getSavedMapping(h) === `SRC:${hA}`); 
        if (memTarget) targetIdx = p.dataB.headers.indexOf(memTarget); 
        else targetIdx = p.dataB.headers.findIndex(h => h.toLowerCase() === hA.toLowerCase()); 
        if (targetIdx !== -1) { 
            p.mapping.push({ name: p.dataB.headers[targetIdx], idxB: targetIdx, targetType: 'source', targetVal: iA }); 
        } 
    }); 
}

function renderDashboard() { 
    const p = projects[activeProjectIdx]; 
    if (!p.mapping || p.mapping.length === 0) { 
        document.getElementById('summaryCards').innerHTML = "<p>Please map columns first.</p>"; 
        return; 
    } 
    const maxRows = Math.max(p.dataA.body.length, p.dataB.body.length); 
    let totalMatches = 0, totalMismatches = 0; 
    const cardArea = document.getElementById('summaryCards'); 
    cardArea.innerHTML = ""; 
    p.mapping.forEach(map => { 
        let match = 0, miss = 0; 
        const lowerName = map.name.toLowerCase(); 
        const isPrice = /price|cost|retail/i.test(lowerName); 
        const isQty = /qty|quantity/i.test(lowerName); 
        for (let i = 0; i < maxRows; i++) { 
            let vB = (p.dataB.body[i]?.[map.idxB] || "").trim(); 
            let vA = ""; 
            if (map.targetType === 'matrix') { 
                vA = p.matrix.find(m => m.key === map.targetVal)?.val || ""; 
            } else { 
                vA = (p.dataA.body[i]?.[map.targetVal] || "").trim(); 
            } 
            let normA = String(vA).trim().toLowerCase(); 
            let normB = String(vB).trim().toLowerCase(); 
            let equal = false; 
            if (isPrice) { 
                equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,'')); 
            } else if (isQty) { 
                equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,'')); 
            } else { 
                equal = (normA === normB); 
            } 
            if (equal) match++; else miss++; 
        } 
        totalMatches += match; totalMismatches += miss; 
        const cls = miss > 0 ? 'bg-warn' : 'bg-good'; 
        cardArea.innerHTML += `<div class="field-card ${cls}"><div class="fc-head">${map.name}</div><div class="fc-stats"><span class="fc-ok">✓ ${match}</span><span class="fc-err">✗ ${miss}</span></div></div>`; 
    }); 
    document.getElementById('globalStats').innerHTML = `<div class="big-stat blue"><div class="bs-val">${maxRows}</div><div class="bs-lbl">Rows</div></div><div class="big-stat blue"><div class="bs-val">${p.mapping.length}</div><div class="bs-lbl">Fields</div></div><div class="big-stat green"><div class="bs-val">${totalMatches}</div><div class="bs-lbl">Matches</div></div><div class="big-stat red"><div class="bs-val">${totalMismatches}</div><div class="bs-lbl">Mismatches</div></div>`; 
    renderResultTables(maxRows); 
}

function renderResultTables(maxRows) { 
    const p = projects[activeProjectIdx]; 
    const tA = document.getElementById('renderTableA'); 
    const tB = document.getElementById('renderTableB'); 
    
    let hA = "<thead><tr><th>#</th>" + p.dataA.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    let hB = "<thead><tr><th>#</th>" + p.dataB.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    let bA = "", bB = ""; 
    const mapLookup = {}; 
    p.mapping.forEach(m => mapLookup[m.idxB] = m); 
    const reverseLookup = {}; 
    p.mapping.forEach(m => { 
        if (m.targetType === 'source') { 
            if (!reverseLookup[m.targetVal]) reverseLookup[m.targetVal] = []; 
            reverseLookup[m.targetVal].push(m); 
        } 
    }); 
    for (let i = 0; i < maxRows; i++) { 
        let rA = `<td>${i+1}</td>`; 
        for (let cA = 0; cA < p.dataA.headers.length; cA++) { 
            let vA = (p.dataA.body[i]?.[cA] || "").trim(); 
            let cls = ""; 
            if (reverseLookup[cA]) { 
                let allMatch = true; 
                reverseLookup[cA].forEach(map => { 
                    let vB = (p.dataB.body[i]?.[map.idxB] || "").trim(); 
                    let normA = String(vA).trim().toLowerCase(); 
                    let normB = String(vB).trim().toLowerCase(); 
                    const lowerName = map.name.toLowerCase(); 
                    const isPrice = /price|cost|retail/i.test(lowerName); 
                    const isQty = /qty|quantity/i.test(lowerName); 
                    let equal = false; 
                    if (isPrice) { 
                        equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,'')); 
                    } else if (isQty) { 
                        equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,'')); 
                    } else { 
                        equal = (normA === normB); 
                    } 
                    if (!equal) allMatch = false; 
                }); 
                cls = allMatch ? "match" : "diff"; 
            } 
            rA += `<td class="${cls}">${vA}</td>`; 
        } 
        let rB = `<td>${i+1}</td>`; 
        p.dataB.headers.forEach((_, colIdx) => { 
            let vB = (p.dataB.body[i]?.[colIdx] || "").trim(); 
            let cls = ""; 
            if (mapLookup.hasOwnProperty(colIdx)) { 
                let map = mapLookup[colIdx]; 
                let vA = ""; 
                if (map.targetType === 'matrix') { 
                    vA = p.matrix.find(m => m.key === map.targetVal)?.val || ""; 
                } else { 
                    vA = (p.dataA.body[i]?.[map.targetVal] || "").trim(); 
                } 
                let normA = String(vA).trim().toLowerCase(); 
                let normB = String(vB).trim().toLowerCase(); 
                const lowerName = p.dataB.headers[colIdx].toLowerCase(); 
                const isPrice = /price|cost|retail/i.test(lowerName); 
                const isQty = /qty|quantity/i.test(lowerName); 
                let equal = false; 
                if (isPrice) { 
                    equal = (normA.replace(/[$,\s]/g,'') === normB.replace(/[$,\s]/g,'')); 
                } else if (isQty) { 
                    equal = (normA.replace(/[\,\s]/g,'') === normB.replace(/[\,\s]/g,'')); 
                } else { 
                    equal = (normA === normB); 
                } 
                if (!equal) { 
                    cls = "diff"; 
                } else { 
                    cls = "match"; 
                } 
            } 
            rB += `<td class="${cls}">${vB}</td>`; 
        }); 
        
        bA += `<tr>${rA}</tr>`; 
        bB += `<tr id="rowB-${i}">${rB}</tr>`; 
    } 
    tA.innerHTML = hA + bA + "</tbody>"; 
    tB.innerHTML = hB + bB + "</tbody>"; 
}

function getSavedMapping(headerName) {
    return localStorage.getItem("map_" + headerName);
}

function saveMappingMemory(headerName, type, val) {
    localStorage.setItem("map_" + headerName, `${type === 'source' ? 'SRC' : 'MAT'}:${val}`);
}
