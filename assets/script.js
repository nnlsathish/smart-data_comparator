// ==========================================
// 1. GLOBAL VARIABLES & INITIALIZATION
// ==========================================

let projects = [];
let activeProjectIdx = 0;
let isOverviewMode = false;
const MAX_TOTAL_SETS = 50; 

// Undo History Variables
let deletedItem = null;      
let deletedItemIdx = -1;     
let deletedRowsHistory = []; 

// --- Window Load ---
window.onload = function() {
    console.log("Validation Tool Loaded Successfully.");
    
    // Initialize first empty project
    addNewProject();

    // Prevent accidental refresh if work is unsaved
    window.addEventListener('beforeunload', function (e) {
        if (hasUnsavedData()) {
            e.preventDefault();
            e.returnValue = 'Unsaved changes'; 
            return 'Unsaved changes';
        }
    });

    // Enter key to close modal
    window.addEventListener('keydown', function(event) {
        const modal = document.getElementById('customModal');
        if (event.key === 'Enter' && modal && modal.classList.contains('open')) {
            event.preventDefault(); 
            document.getElementById('modalBtn').click(); 
        }
    });
};

// Check if there is data to lose
function hasUnsavedData() {
    const inputA = document.getElementById('tableA');
    const inputB = document.getElementById('tableB');
    
    // Check text areas
    const hasVisibleData = (inputA?.value.trim().length > 0) || (inputB?.value.trim().length > 0);
    
    // Check project memory
    const hasStoredData = projects.some(p => p.status !== 'empty');
    
    return hasVisibleData || hasStoredData;
}

// ==========================================
// 2. APP MANAGEMENT (Reset, Add Sets)
// ==========================================

function hardResetApp() {
    if (!confirm("Start Fresh? This will delete all current sets and data.")) {
        return;
    }

    // 1. Reset Global Variables
    projects = [];
    activeProjectIdx = 0;
    isOverviewMode = false;
    deletedItem = null;
    deletedItemIdx = -1;
    deletedRowsHistory = [];

    // 2. Clear File Inputs (Hidden ones)
    document.querySelectorAll('input[type="file"]').forEach(el => el.value = "");
    
    // 3. Reset UI State (Sidebar & Tabs)
    document.querySelector('.sidebar').classList.remove('overview-locked');
    document.getElementById('tabOverview').classList.remove('active');
    document.getElementById('overviewSection').style.display = 'none';

    // 4. Reset to Step 1 visually
    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById('step1').style.display = 'block';
    
    // Reset Sidebar Navigation Styles
    document.querySelectorAll('.v-step').forEach(el => el.classList.remove('active'));
    document.getElementById('navStep1').classList.add('active');

    // 5. Initialize a fresh Project
    createSet(); // Creates "Set 1"
    renderTopBar();
    
    // 6. Load the empty project (This clears the textareas and tables automatically)
    loadProjectIntoView(0);

    // 7. Feedback
    showToast("Application Reset Successfully");
}

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
    document.getElementById(`headerMode${side}`).value = mode;
    const group = document.getElementById(`groupMode${side}`);
    group.querySelectorAll('.toggle-btn').forEach(btn => {
        if(btn.getAttribute('data-val') === mode) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

function addSetsFromInput() {
    const input = document.getElementById('addSetQty');
    let qty = parseInt(input.value) || 1;
    
    if (qty < 1) qty = 1;
    
    if (projects.length + qty > MAX_TOTAL_SETS) {
        showModal("Limit Reached", `Cannot add more than ${MAX_TOTAL_SETS} sets.`, 'error');
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

function createSet(customName = null, groupFile = null) {
    const id = projects.length + 1;
    projects.push({
        name: customName || `Set ${id}`,
        fileName: groupFile || "Manual Input", // Store the file name for grouping
        status: 'empty', 
        rawA: "", rawB: "", rawMatrix: "",
        dataA: null, dataB: null, matrix: [], mapping: [], 
        step: 1, showMatrix: false,
        summary: { matches: 0, mismatches: 0 }
    });
}

// ==========================================
// 3. FILE UPLOAD HANDLERS
// ==========================================

function handleBulkUpload(input) {
    if (!input.files || input.files.length === 0) return;

    // --- FIX: Save current state first so we don't lose manual edits ---
    saveCurrentViewToProject(); 
    
    const files = Array.from(input.files);
    let totalCreated = 0;
    
    // Updated Safe Check: Check BOTH tables before wiping
    // We only wipe if we are on the first empty set AND both tables are visually empty
    const currentA = document.getElementById('tableA').value.trim();
    const currentB = document.getElementById('tableB').value.trim(); 

    if (projects.length === 1 && projects[0].status === 'empty' && !projects[0].rawA && !currentA && !currentB) {
        projects = []; 
        activeProjectIdx = -1;
    }

    const processNextFile = (fileIdx) => {
        if (fileIdx >= files.length) {
            if (input.value) input.value = ""; 
            
            if (totalCreated > 0) {
                renderTopBar();
                showModal("Import Successful", `Loaded <strong>${totalCreated}</strong> sheet(s).`, "success");
                
                // If we wiped the initial empty set, switch to 0. Otherwise switch to the first NEW set.
                let targetIdx = (activeProjectIdx === -1) ? 0 : (projects.length - totalCreated);
                if (targetIdx < 0) targetIdx = 0;
                
                activeProjectIdx = targetIdx;
                switchProject(targetIdx, true); 
            } else {
                showModal("No Data Found", "Sheets were empty, hidden, or had 0 Quantity.", "error");
            }
            return;
        }

        const file = files[fileIdx];
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellStyles: true });

                workbook.SheetNames.forEach((sheetName, index) => {
                    // --- 1. IGNORE HIDDEN SHEETS ---
                    if (workbook.Workbook && workbook.Workbook.Sheets) {
                        const sMeta = workbook.Workbook.Sheets.find(s => s.name === sheetName);
                        if (sMeta && (sMeta.Hidden !== 0 || sMeta.state === 'hidden')) {
                            return; // Skip this sheet
                        }
                    }

                    const sheet = workbook.Sheets[sheetName];
                    let rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false }); 

                    // --- 2. IGNORE HIDDEN COLUMNS ---
                    if (sheet['!cols']) {
                        const hiddenIndices = new Set();
                        sheet['!cols'].forEach((col, i) => {
                            if (!col) return;
                            const isExplicitHidden = col.hidden === true || col.hidden === 1;
                            const isZeroWidth = (col.wpx != null && col.wpx < 1) || (col.width != null && col.width < 0.1);
                            
                            if (isExplicitHidden || isZeroWidth) {
                                hiddenIndices.add(i);
                            }
                        });

                        if (hiddenIndices.size > 0) {
                            rawData = rawData.map(row => row.filter((_, i) => !hiddenIndices.has(i)));
                        }
                    }

                    if (rawData.length > 0) {
                        const headerIdx = findHeaderRowIndex(rawData);
                        const cleanTable = extractTableStrict(rawData);
                        
                        // Extract Matrix Data
                        let matrixData = [];
                        if (headerIdx > 0) {
                            matrixData = extractMatrixData(rawData, headerIdx);
                        }
                        
                        // --- 3. SMART FILTER (Check Qty) ---
                        if (cleanTable.length > 1) {
                            const header = cleanTable[0];
                            const qtyIndices = header
                                .map((h, i) => /qty|quantity|units|count|pcs|bill|ship/i.test(String(h).toLowerCase()) ? i : -1)
                                .filter(index => index !== -1);
                            
                            let hasData = false;
                            
                            if (qtyIndices.length > 0) {
                                for (let r = 1; r < cleanTable.length; r++) {
                                    const row = cleanTable[r];
                                    for (let colIdx of qtyIndices) {
                                        const val = row[colIdx];
                                        if (!val) continue;
                                        const num = parseFloat(String(val).replace(/[, ]/g, ''));
                                        if (!isNaN(num) && num > 0) {
                                            hasData = true;
                                            break; 
                                        }
                                    }
                                    if (hasData) break;
                                }
                            } else {
                                hasData = true; // Fallback
                            }
                            
                            // --- 4. CREATE SET ---
                            if (hasData) {
                                totalCreated++;
                                createSet(sheetName, file.name); 
                                
                                let p = projects[projects.length - 1];
                                p.rawA = arrayToTSV(cleanTable);
                                
                                if (matrixData.length > 0) {
                                    const cleanList = matrixData.map(m => {
                                        let k = m.key.replace(/[\r\n]+/g, " ").trim();
                                        k = k.replace(/"/g, "").replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim();
                                        
                                        const upperK = k.toUpperCase();
                                        if (upperK.includes("ADAPTIVE")) k = "ADAPTIVE";
                                        else if (upperK.includes("UPF")) k = "UPF";
                                        else if (upperK.includes("LYCRA")) k = "LYCRA";
                                        else if (upperK.includes("SUSTAINABILITY")) k = "SUSTAINABILITY";
                                        else if (upperK.includes("SKU")) k = "SKU#";
                                        else if (upperK.includes("GOTS INFO")) k = "GOTS INFO";
                                        else if (upperK.includes("GOTS ICON")) k = "GOTS Icon";

                                        return `${k}: ${m.val}`;
                                    });

                                    p.rawMatrix = cleanList.join("\n");
                                    p.matrix = parseMatrixString(p.rawMatrix);
                                    p.showMatrix = true; 
                                }

                                p.status = 'ready';
                            }
                        }
                    }
                });

            } catch (err) {
                console.error("Error reading file:", file.name, err);
            } finally {
                processNextFile(fileIdx + 1);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    processNextFile(0);
}

// ==========================================
// 4. DATA EXTRACTION LOGIC
// ==========================================

function extractTableStrict(data) {
    let startIndex = -1;
    const qtyRegex = /qty|quantity/i;
    const chineseRegex = /[\u4e00-\u9fff]/; 
    let searchStartRow = 0;

    for (let i = 0; i < Math.min(data.length, 30); i++) {
        if (!data[i]) continue;
        const rowText = data[i].join(" ").toLowerCase();
        if (rowText.includes("sku information")) {
            searchStartRow = i + 1; 
            break;
        }
    }

    for (let i = searchStartRow; i < Math.min(data.length, 300); i++) {
        let row = data[i];
        if (!row) continue;
        const rowText = row.join(" ").toLowerCase();
        
        if (rowText.includes("tel#") || rowText.includes("email:")) continue; 
        if (rowText.includes("just fill total qty") || rowText.includes("no moq")) continue;
        if (rowText.includes("consider wastage")) continue;
        if (rowText.includes("refer to the chart")) continue;
        if (rowText.includes("kohls po quantities")) continue;
        if (rowText.includes("overrun") && rowText.includes("ordering qty")) continue;
        
        const hasQty = row.some(cell => cell && qtyRegex.test(String(cell).trim()));
        const hasChinese = row.some(cell => cell && chineseRegex.test(String(cell).trim()));
        
        if (hasQty && hasChinese) { startIndex = i; break; }
        if (hasQty && !rowText.includes("round up")) { startIndex = i; break; }
        if (rowText.includes("description10") || rowText.includes("description 10")) { startIndex = i; break; }
    }

    if (startIndex === -1) return [];

    let headerRow = data[startIndex];
    let nextRow = data[startIndex + 1];
    let dataStartIndex = startIndex + 1; 

    if (nextRow && nextRow.join(" ").toLowerCase().includes("#n/a")) {
        nextRow = data[startIndex + 2];
        dataStartIndex = startIndex + 2; 
    }

    let useCombinedHeader = false;
    if (nextRow) {
        const subKeywords = ["UPC", "EAN", "FEATURE ICON"];
        useCombinedHeader = nextRow.some(c => c && subKeywords.some(k => String(c).toUpperCase().includes(k)));
    }

    let headerStartCol = headerRow.findIndex(c => String(c).trim() !== "");
    if (headerStartCol === -1) headerStartCol = 0;

    let fixedTableWidth = null;
    let desc10Index = headerRow.findIndex(c => c && String(c).toLowerCase().replace(/\s/g, "") === "description10");
    if (desc10Index !== -1 && desc10Index >= headerStartCol) {
        fixedTableWidth = desc10Index - headerStartCol + 1; 
    }

    let cleanRows = [];
    let finalHeader = [];
    
    if (useCombinedHeader) {
        let combinedHeader = [];
        const maxLen = Math.max(headerRow.length, nextRow.length);
        for (let c = headerStartCol; c < maxLen; c++) {
            let val1 = headerRow[c] ? String(headerRow[c]).trim() : "";
            let val2 = nextRow[c] ? String(nextRow[c]).trim() : "";
            combinedHeader.push((val2 && val2.length > 1) ? val2 : val1);
        }
        if(fixedTableWidth) combinedHeader = combinedHeader.slice(0, fixedTableWidth);
        finalHeader = combinedHeader;
    } else {
        let extractedHeader = headerRow.slice(headerStartCol);
        if(fixedTableWidth) extractedHeader = extractedHeader.slice(0, fixedTableWidth);
        finalHeader = extractedHeader;
    }

    for (let k = 1; k < finalHeader.length; k++) {
        let current = finalHeader[k];
        let prev = finalHeader[k-1];
        if ((!current || current.toString().trim() === "") && prev && prev.toString().trim() !== "") {
            finalHeader[k] = prev; 
        }
    }

    cleanRows.push(finalHeader);
    let emptyRowCount = 0;

    for (let i = dataStartIndex; i < data.length; i++) {
        let row = data[i];
        if (!row || row.every(c => !c || String(c).trim() === "")) { 
            emptyRowCount++; 
            if (emptyRowCount >= 10) break; 
            continue; 
        }
        const rowStr = row.join(" ").toLowerCase();
        
        const currentRowHasQty = row.some(cell => cell && qtyRegex.test(String(cell).trim()));
        const currentRowHasChinese = row.some(cell => cell && chineseRegex.test(String(cell).trim()));
        if (currentRowHasQty && currentRowHasChinese) break; 
        
        if (rowStr.includes("#n/a") || rowStr.includes("#ref!")) continue;
        if (rowStr.includes("(max") && rowStr.includes("digits)")) continue; 
        if (rowStr.includes("remark") && rowStr.includes("max length")) continue;

        const firstTextIndex = row.findIndex(c => c && String(c).trim().length > 0);
        let firstText = "";
        if (firstTextIndex !== -1) firstText = String(row[firstTextIndex]).toLowerCase().trim();
        if (firstText === "eg" || firstText.startsWith("eg ") || firstText.includes("e.g.")) continue;
        
        if (firstText.includes("disclaimer") || firstText.startsWith("total") || 
            firstText.startsWith("notes") || firstText.startsWith("remarks") || 
            rowStr.includes("page") && rowStr.includes("of")) break; 
        
        emptyRowCount = 0; 
        let alignedRow = row.slice(headerStartCol);
        if (fixedTableWidth) alignedRow = alignedRow.slice(0, fixedTableWidth);
        cleanRows.push(alignedRow);
    }

    let header = cleanRows[0];
    if (header && cleanRows.length > 1) {
        let body = cleanRows.slice(1);
        for (let i = 0; i < header.length; i++) {
            let currentHeader = header[i];
            if (!currentHeader || currentHeader.trim() === "") continue;

            let j = i + 1;
            while (j < header.length && header[j] === currentHeader) j++;

            if (j > i + 1) { 
                let densityMap = [];
                for (let k = i; k < j; k++) {
                    let count = 0;
                    for (let r = 0; r < body.length; r++) {
                        if (body[r][k] && body[r][k].toString().trim().length > 0) count++;
                    }
                    densityMap.push({ index: k, count: count });
                }
                
                densityMap.sort((a, b) => b.count - a.count);
                let winner = densityMap[0];
                let indexToKeep = (winner.count > 0) ? winner.index : i;

                for (let k = i; k < j; k++) {
                    if (k !== indexToKeep) header[k] = "";
                }
                i = j - 1; 
            }
        }
    }

    return cleanRows;
}

function arrayToTSV(data) {
    return data.map(row => 
        row.map(cell => (cell == null ? "" : String(cell).replace(/[\r\n]+/g, " ").trim()))
           .join("\t")
    ).join("\n");
}

// ==========================================
// 5. UNDO / RESTORE LOGIC
// ==========================================

function deleteProject(e, idx) {
    e.stopPropagation(); 
    deletedItem = projects[idx];
    deletedItemIdx = idx;
    
    projects.splice(idx, 1);
    
    if (projects.length === 0) { 
        createSet(); 
        activeProjectIdx = 0; 
    } else if (activeProjectIdx >= projects.length) {
        activeProjectIdx = projects.length - 1;
    }
    
    renderTopBar();
    
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
    
    if (deletedItemIdx >= 0 && deletedItemIdx <= projects.length) {
        projects.splice(deletedItemIdx, 0, deletedItem);
    } else {
        projects.push(deletedItem);
    }
    
    deletedItem = null;
    document.getElementById('btnRestoreSet').style.display = 'none';
    renderTopBar();
    
    if(isOverviewMode) showOverview();
    else switchProject(projects.indexOf(deletedItem));
}

// ==========================================
// 6. UI MODALS & ALERTS
// ==========================================

function showModal(title, content, type = 'success', callback = null) {
    const modal = document.getElementById('customModal');
    const titleEl = document.getElementById('modalTitle');
    const msgEl = document.getElementById('modalMsg');
    const iconEl = document.getElementById('modalIcon');
    const btnEl = document.getElementById('modalBtn');
    
    if (callback) {
        btnEl.onclick = function() { callback(); closeModal(); };
    } else {
        btnEl.onclick = closeModal;
    }
    
    if (modal) {
        titleEl.innerText = title;
        msgEl.innerHTML = content;
        
        if (type === 'error') {
            iconEl.className = 'fas fa-times-circle sa-icon-error';
            iconEl.style.color = '#ef4444';
            btnEl.style.backgroundColor = '#ef4444';
            btnEl.innerText = 'Close';
        } else if (type === 'confirm') {
            iconEl.className = 'fas fa-question-circle sa-icon-warn';
            iconEl.style.color = '#f59e0b';
            btnEl.style.backgroundColor = '#f59e0b';
            btnEl.innerText = 'Confirm';
        } else {
            iconEl.className = 'fas fa-check-circle sa-icon-check';
            iconEl.style.color = '#10b981';
            btnEl.style.backgroundColor = '#2563eb';
            btnEl.innerText = 'OK';
        }
        modal.classList.add('open');
    } else {
        alert(title + "\n" + content.replace(/<br>/g,'\n'));
    }
}

function closeModal() { 
    document.getElementById('customModal')?.classList.remove('open'); 
}

window.onclick = function(event) { 
    if (event.target === document.getElementById('customModal')) closeModal(); 
}

// ==========================================
// 7. TOP BAR & OVERVIEW
// ==========================================

function renderTopBar() {
    const list = document.getElementById('projectTabs');
    list.innerHTML = "";
    
    let lastFileName = null;

    projects.forEach((p, i) => {
        if (p.fileName !== lastFileName) {
            const shortName = (p.fileName.length > 15) ? p.fileName.substring(0, 12) + "..." : p.fileName;
            list.innerHTML += `<div style="padding: 0 8px; display:flex; align-items:center; font-size:10px; color:#999; border-left:1px solid #ddd; margin-left:4px;">${shortName}</div>`;
            lastFileName = p.fileName;
        }

        const activeClass = (i === activeProjectIdx && !isOverviewMode) ? 'active' : '';
        list.innerHTML += `
            <div class="tab-item ${activeClass}" onclick="switchProject(${i})">
                <div class="tab-dot ${p.status}"></div>
                <span>${p.name}</span>
                <button class="btn-tab-close" onclick="deleteProject(event, ${i})">×</button>
            </div>`;
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
    
    let totalSets = projects.length;
    let totalRows = 0; 
    let totalMatches = 0; 
    let totalMismatches = 0; 
    let tbody = "";

    projects.forEach(p => {
        const matches = p.summary ? p.summary.matches : 0;
        const mismatches = p.summary ? p.summary.mismatches : 0;
        totalRows += (p.dataA ? p.dataA.body.length : 0);
        totalMatches += matches; 
        totalMismatches += mismatches;
        
        tbody += `<tr>
                <td><strong>${p.name}</strong></td>
                <td><span class="status-dot ${p.status}"></span> ${p.status.toUpperCase()}</td>
                <td>${p.dataA ? p.dataA.body.length : 0}</td>
                <td>${p.dataB ? p.dataB.body.length : 0}</td>
                <td style="color:#10b981; font-weight:bold;">${matches}</td>
                <td style="color:#ef4444; font-weight:bold;">${mismatches}</td>
                <td><button onclick="switchProject(${projects.indexOf(p)})" class="btn-ghost">View</button></td>
            </tr>`;
    });

    document.getElementById('overviewBody').innerHTML = tbody;
    document.getElementById('overviewStats').innerHTML = `
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${totalSets}</div><div class="bs-lbl">Sets</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${totalRows}</div><div class="bs-lbl">Rows</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#10b981">${totalMatches}</div><div class="bs-lbl">Matches</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#ef4444">${totalMismatches}</div><div class="bs-lbl">Mismatches</div></div>`;
}

function switchProject(idx, skipSave = false) {
    if(isOverviewMode) {
        document.getElementById('tabOverview').classList.remove('active');
        document.querySelector('.sidebar').classList.remove('overview-locked');
        isOverviewMode = false;
    } else if (!skipSave) { 
        saveCurrentViewToProject(); 
    }
    
    activeProjectIdx = idx;
    deletedRowsHistory = [];
    document.getElementById('btnUndoRow').style.display = 'none';
    
    loadProjectIntoView(idx);
    renderTopBar();
}

function saveCurrentViewToProject() {
    if (projects.length === 0 || isOverviewMode || activeProjectIdx < 0) return;
    
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


// =========================================
// MATRIX EXTRACTION & UI FUNCTIONS (RESTORED)
// =========================================

// 1. Scan ABOVE the table for Matrix Keys
function extractMatrixData(data, endRow) {
    const matrix = [];
    // KEYWORDS MUST BE UPPERCASE to match correctly
    const keywords = [
        "ADAPTIVE", "UPF", "LYCRA", "SUSTAINABILITY",
        "SKU", "DEPT", "MAJOR CLASS", "SUB CLASS",
        "DESCRIPTION", "OPTIONS", 
        "GOTS INFO", "GOTS ICON" 
    ];

    for (let i = 0; i < endRow; i++) {
        const row = data[i];
        if (!row) continue;

        for (let c = 0; c < row.length; c++) {
            const cell = String(row[c] || "").trim();
            if (!cell) continue;

            const upperCell = cell.toUpperCase();

            // --- IGNORE RULES ---
            if (upperCell.includes("INSTRUCTION")) continue;
            if (upperCell.includes("FEATURE OPTION")) continue;
            if (upperCell.includes("FEATURES OPTION")) continue;
            if (upperCell.includes("DRAW FROM BULK")) continue;
            // --------------------

            const matchedKeyword = keywords.find(k => upperCell.includes(k));

            if (matchedKeyword) {
                let candidateVal = "";
                let foundVal = false;

                // 1. Look Right
                if (row[c+1] && String(row[c+1]).trim() !== "") {
                    candidateVal = String(row[c+1]).trim();
                    foundVal = true;
                } 
                // 2. Look Skip-Right (merged cells)
                else if (row[c+2] && String(row[c+2]).trim() !== "") {
                    candidateVal = String(row[c+2]).trim();
                    foundVal = true;
                }

                // CHECK: Is the "Value" actually just another "Label"?
                let isNextLabel = false;
                if (foundVal) {
                    const hasKeyword = keywords.some(k => candidateVal.toUpperCase().includes(k));
                    if (hasKeyword) {
                        if (candidateVal.includes("(") || candidateVal.includes("\n") || candidateVal.toUpperCase().includes("INSTRUCTION")) {
                            isNextLabel = true;
                        }
                    }
                }

                const finalVal = isNextLabel ? "" : candidateVal;
                const cleanKey = cell.replace(/[\r\n]+/g, " ").trim();
                
                // Prevent duplicates
                if (!matrix.some(m => m.key === cleanKey)) {
                    matrix.push({ key: cleanKey, val: finalVal });
                }
            }
        }
    }
    return matrix;
}

// 2. Parse the text from the Input Box
function parseMatrixString(rawString) {
    if (!rawString) return [];
    const rows = [];
    const rawLines = rawString.split(/\n/); 
    
    rawLines.forEach(line => {
        // SAFETY FIX: Remove parentheses content BEFORE finding the colon
        let cleanLine = line.replace(/\(.*?\)/g, "").trim(); 

        let k = "", v = "";
        
        if (cleanLine.includes(":")) {
            let idx = cleanLine.indexOf(":");
            k = cleanLine.substring(0, idx);
            v = cleanLine.substring(idx+1);
        } else {
            k = cleanLine; 
        }

        // Cleanup
        k = k.replace(/"/g, "");
        k = k.replace(/\s+/g, " ").trim();
        v = v ? v.trim() : "";

        // Standardize
        const upperK = k.toUpperCase();
        if (upperK.includes("ADAPTIVE")) k = "ADAPTIVE";
        else if (upperK.includes("UPF")) k = "UPF";
        else if (upperK.includes("LYCRA")) k = "LYCRA";
        else if (upperK.includes("SUSTAINABILITY")) k = "SUSTAINABILITY";
        else if (upperK.includes("SKU")) k = "SKU#";
        else if (upperK.includes("GOTS INFO")) k = "GOTS INFO";
        else if (upperK.includes("GOTS ICON")) k = "GOTS Icon";
        
        if (k) rows.push({ key: k, val: v });
    });
    return rows;
}

// 3. UI Helpers
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
        if (row.includes("\t")) {
            [k, v] = row.split("\t");
        } else if (row.includes(":")) {
            let idx = row.indexOf(":");
            k = row.substring(0, idx);
            v = row.substring(idx+1);
        } else if (row.includes("=")) { 
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

// 4. String Cleaners
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

function loadProjectIntoView(idx) {
    const p = projects[idx];
    if(!p) return; 

    document.getElementById('tableA').value = p.rawA || "";
    document.getElementById('tableB').value = p.rawB || "";
    
    document.getElementById('previewTableA').innerHTML = "";
    document.getElementById('previewTableB').innerHTML = "";
    document.getElementById('countA').innerText = "0";
    document.getElementById('countB').innerText = "0";
    document.getElementById('mappingBody').innerHTML = "";
    
    document.getElementById('matrixRawInput').value = p.rawMatrix || "";
    const matrixList = document.getElementById('matrixList');
    if (matrixList) {
        matrixList.innerHTML = "";
        if (p.matrix && p.matrix.length > 0) {
            p.matrix.forEach(m => addMatrixRow(m.key, m.val));
        } else {
            addMatrixRow(); 
        }
    }
    
    const matSec = document.getElementById('matrixSection');
    const btn = document.getElementById('btnToggleMatrix');
    if (matSec && btn) {
        matSec.style.display = p.showMatrix ? 'block' : 'none';
        btn.innerHTML = p.showMatrix 
            ? `<i class="fas fa-minus-circle"></i> Hide Matrix Rules` 
            : `<i class="fas fa-plus-circle"></i> Show Matrix Rules (Optional)`;
    }
    
    if (p.step === 3) renderMappingTable(); 
    jumpToStep(p.step || 1);
    
    if (p.step === 4) renderDashboard();
    else if (p.step === 2 && p.dataA) renderPreviewTables();
}

function clearCurrentProject() {
    const p = projects[activeProjectIdx];
    p.rawA = ""; 
    p.rawB = ""; 
    p.rawMatrix = "";
    p.dataA = null; 
    p.dataB = null; 
    p.matrix = []; 
    p.mapping = [];
    p.status = 'empty'; 
    p.step = 1; 
    p.showMatrix = false;
    p.summary = { matches: 0, mismatches: 0 };
    
    loadProjectIntoView(activeProjectIdx); 
    renderTopBar();
}

function jumpToStep(step) {
    const p = projects[activeProjectIdx];
    p.step = step;
    
    if(step === 3) renderMappingTable();
    if(step === 4) renderDashboard();

    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById(`step${step}`).style.display = 'block';
    
    document.querySelectorAll('.v-step').forEach(el => el.classList.remove('active'));
    document.getElementById(`navStep${step}`).classList.add('active');
}

// ==========================================
// 8. DATA IMPORT (Sainpase Excel)
// ==========================================

function handleCPQUpload(input) {
    if (!input.files || input.files.length === 0) return;

    // --- FIX: Save Table 1 data before processing Table 2 upload ---
    saveCurrentViewToProject();

    const files = Array.from(input.files).sort((a, b) => a.lastModified - b.lastModified);
    
    let processedCount = 0;
    const startIdx = activeProjectIdx;

    const processNext = (i) => {
        if (i >= files.length) {
            input.value = ""; 
            showModal("Upload Complete", `Successfully loaded <strong>${processedCount}</strong> files (Sorted by Download Time).`, "success");
            loadProjectIntoView(activeProjectIdx);
            return;
        }

        const targetIdx = startIdx + i;
        
        if (targetIdx >= projects.length) {
            input.value = ""; 
            showModal("Upload Partial", `Loaded ${processedCount} files.<br>Stopped because there are no more Sets to fill.`, "warning");
            loadProjectIntoView(activeProjectIdx);
            return;
        }

        const file = files[i];
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                let extracted = "";
                
                if (workbook.SheetNames.length > 1) {
                    const sheet2 = workbook.Sheets[workbook.SheetNames[1]];
                    const rows = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
                    let start = -1;
                    for (let r = 0; r < rows.length; r++) {
                        if ((rows[r]||[]).join(" ").toLowerCase().includes("sku information")) { 
                            start = r + 1; 
                            break; 
                        }
                    }
                    if(start !== -1) {
                        extracted = rows.slice(start)
                            .map(r => r.slice(1).map(c => (c==null)?"":c).join("\t"))
                            .join("\n");
                    }
                }
                
                if(!extracted) {
                    const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
                    extracted = rows.map(r => r.map(c => (c==null)?"":c).join("\t")).join("\n");
                }

                projects[targetIdx].rawB = extracted;
                projects[targetIdx].dataB = null; 
                processedCount++;
                
            } catch (err) {
                console.error("Error reading file for Set " + (targetIdx+1), err);
            } finally {
                processNext(i + 1);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    processNext(0);
}

// ==========================================
// 9. PREVIEW & CLEANING
// ==========================================
function goToPreview() {
    // --- FIX START: Get the latest data from the text boxes ---
    const inputA = document.getElementById('tableA');
    const inputB = document.getElementById('tableB');
    
    // Update the project memory with your manual changes immediately
    const p = projects[activeProjectIdx];
    p.rawA = inputA.value; 
    p.rawB = inputB.value;
    // --- FIX END ---

    // 3. Validation
    if (!p.rawA.trim() || !p.rawB.trim()) {
        showModal("Missing Data", "Please paste data into both tables.", 'error');
        return;
    }
    
    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";
    
    // 4. Parse the NEW data (your edits)
    p.dataA = parseExcelData(p.rawA, modeA);
    p.dataB = parseExcelData(p.rawB, modeB);
    
    if (!p.dataA || !p.dataB) {
        showModal("Error Parsing", "Could not parse data.", 'error');
        return;
    }
    
    // 5. Clean the NEW data
    autoCleanData(p.dataA);
    autoCleanData(p.dataB);
    
    // 6. Move to next step
    p.status = 'ready'; 
    p.step = 2;
    renderTopBar(); 
    renderPreviewTables(); 
    jumpToStep(2);
}

function renderPreviewTables() {
    const p = projects[activeProjectIdx];
    
    if (p.dataA) {
        renderSinglePreview('previewTableA', 'countA', p.dataA, 'A');
    } else {
        document.getElementById('previewTableA').innerHTML = "";
        document.getElementById('countA').innerText = "0";
    }

    if (p.dataB) {
        renderSinglePreview('previewTableB', 'countB', p.dataB, 'B');
    } else {
        document.getElementById('previewTableB').innerHTML = "";
        document.getElementById('countB').innerText = "0";
    }
    
    const undoBtn = document.getElementById('btnUndoRow');
    if(deletedRowsHistory.length > 0) {
        undoBtn.style.display = 'inline-flex';
        undoBtn.innerText = `Undo Delete (${deletedRowsHistory.length})`;
    } else {
        undoBtn.style.display = 'none';
    }
}

function renderSinglePreview(containerId, countId, data, side) {
    if (!data || !data.body) return;

    document.getElementById(countId).innerText = data.body.length;

    const container = document.getElementById(containerId);
    const tableBox = container.closest('.table-box');
    const header = tableBox.querySelector('.box-header');

    let toolsDiv = header.querySelector('.custom-header-tools');
    
    if (!toolsDiv) {
        if (header.style.justifyContent !== 'space-between') {
            header.style.display = 'flex';
            header.style.justifyContent = 'space-between';
            header.style.alignItems = 'center';
        }

        toolsDiv = document.createElement('div');
        toolsDiv.className = 'custom-header-tools';
        toolsDiv.style.display = 'flex';
        toolsDiv.style.gap = '5px';
        toolsDiv.style.alignItems = 'center';

        toolsDiv.innerHTML = `
            <input type="text" id="rangeInput${side}" placeholder="e.g. 2-5" 
                style="padding:4px 8px; border-radius:4px; border:none; font-size:12px; color:#333; width:80px; outline:none;" 
                onclick="event.stopPropagation()">
            <button onclick="deleteRange('${side}'); event.stopPropagation()" 
                style="background:rgba(255,255,255,0.2); color:white; border:1px solid rgba(255,255,255,0.4); 
                       padding:4px 10px; border-radius:4px; cursor:pointer; font-size:11px; font-weight:600;">
                <i class="fas fa-trash-alt"></i> Range
            </button>
        `;
        header.appendChild(toolsDiv);
    }
    
    let html = `
    <table class="clean-table" id="previewTable${side}">
        <thead>
            <tr>
                <th style="width:50px; text-align:center;">Action</th>
                <th>#</th>`;
    
    data.headers.forEach(h => html += `<th>${h}</th>`);
    
    html += "</tr></thead><tbody>";
    
    data.body.forEach((row, idx) => {
        html += `<tr>
            <td style="text-align:center;">
                <button class="btn-ghost" style="color:#ef4444; font-size:12px;" onclick="deleteRow('${side}', ${idx})">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
            <td>${idx+1}</td>`;
        
        row.forEach(cell => html += `<td>${cell}</td>`);
        html += `</tr>`;
    });
    
    html += "</tbody></table>"; 
    
    container.innerHTML = html;
}

function deleteRange(side) {
    const input = document.getElementById(`rangeInput${side}`);
    const val = input.value.trim();
    if (!val) return;
    
    const parts = val.split('-');
    if (parts.length !== 2) { 
        showModal("Invalid Format", "Use Start-End (e.g. 4-7)", "error"); 
        return; 
    }
    
    let start = parseInt(parts[0]);
    let end = parseInt(parts[1]);
    
    if (isNaN(start) || isNaN(end) || start > end) return;
    
    const p = projects[activeProjectIdx];
    const targetBody = (side === 'A') ? p.dataA.body : p.dataB.body;
    let deleteCount = 0;
    
    for (let i = end; i >= start; i--) {
        const idx = i - 1; 
        if (idx >= 0 && idx < targetBody.length) {
            deletedRowsHistory.push({ side: side, idx: idx, data: targetBody[idx] });
            targetBody.splice(idx, 1);
            deleteCount++;
        }
    }
    
    if (deleteCount > 0) { 
        showToast(`Deleted ${deleteCount} rows.`); 
        renderPreviewTables(); 
    }
}

function deleteRow(side, rowIndex) {
    const p = projects[activeProjectIdx];
    let rowData = (side === 'A') ? p.dataA.body[rowIndex] : p.dataB.body[rowIndex];
    
    if(side === 'A') p.dataA.body.splice(rowIndex, 1); 
    else p.dataB.body.splice(rowIndex, 1);
    
    deletedRowsHistory.push({ side: side, idx: rowIndex, data: rowData });
    renderPreviewTables();
}

function restoreLastRow() {
    if (deletedRowsHistory.length === 0) return;
    
    const last = deletedRowsHistory.pop();
    const p = projects[activeProjectIdx];
    
    if (last.side === 'A') p.dataA.body.splice(last.idx, 0, last.data);
    else p.dataB.body.splice(last.idx, 0, last.data);
    
    renderPreviewTables();
}

// ==========================================
// 10. MAPPING LOGIC
// ==========================================

function goToMapping() {
    const p = projects[activeProjectIdx];
    if (p.dataA.body.length !== p.dataB.body.length) {
        showModal("Row Mismatch", `Form Data: ${p.dataA.body.length} rows<br>Sainpase Data: ${p.dataB.body.length} rows<br><br>Row counts must be equal.`, 'error');
        return; 
    }
    renderMappingTable();
    jumpToStep(3);
}

function renderMappingTable() {
    const p = projects[activeProjectIdx];
    if (!p.dataA || !p.dataB) return false;
    
    const tbody = document.getElementById('mappingBody');
    tbody.innerHTML = "";
    document.getElementById('mapSearchInput').value = "";

    const leftRows = [];
    p.dataA.headers.forEach((h, i) => { 
        if (h && h.trim() !== "") leftRows.push({ type: 'source', name: h, val: i }); 
    });
    p.matrix.forEach(m => leftRows.push({ type: 'matrix', name: m.key, val: m.key, display: m.val }));

    const rightOptions = p.dataB.headers
        .map((h, i) => ({ name: h, index: i }))
        .filter(opt => opt.name && opt.name.trim() !== "" && !opt.name.toLowerCase().startsWith("unnamed"));
    rightOptions.sort((a,b) => a.name.localeCompare(b.name));

    function getStandardKey(headerName) {
        let clean = headerName.toLowerCase()
            .replace(/\(?\boptional\b\)?/g, "")
            .replace(/[\u4e00-\u9fff]/g, "")
            .replace(/[^a-z0-9]/g, "")
            .trim();
            
        if (["quantity", "totalqty", "qty", "billshipquantity"].includes(clean)) return "QTY_GROUP";
        if (["retailprice", "price", "retail", "retail1"].includes(clean)) return "PRICE_GROUP";
        if (clean.includes("barcode") || ["upc", "gtin"].includes(clean)) return "BARCODE_GROUP";
        if (["subclass", "sku"].includes(clean)) return "SUB_CLASS_GROUP";
        if (["majorclass", "majclass", "class", "major"].includes(clean)) return "MAJOR_CLASS_GROUP";
        if (["department", "dept"].includes(clean)) return "DEPT_GROUP";
        if (["color", "colour", "clr"].includes(clean)) return "COLOR_GROUP";
        return clean;
    }

    leftRows.forEach(row => {
        let matchVal = "-1";
        
        const savedSession = p.mapping.find(m => {
            return (row.type === 'source' && m.targetType === 'source' && m.targetVal === row.val) ||
                   (row.type === 'matrix' && m.targetType === 'matrix' && m.targetVal === row.val);
        });

        if (savedSession) {
            matchVal = savedSession.idxB;
        } else {
            const cleanSource = row.name.toLowerCase().trim();
            
            let directMatch = rightOptions.find(opt => opt.name.toLowerCase().trim() === cleanSource);
            
            if (!directMatch) {
                 directMatch = rightOptions.find(opt => {
                     const t = opt.name.toLowerCase().trim();
                     return t === cleanSource || t === cleanSource + " name" || t === cleanSource + " id";
                 });
            }

            if (directMatch) {
                matchVal = directMatch.index;
            } else {
                const sourceGroup = getStandardKey(row.name);
                const groupMatch = rightOptions.find(target => getStandardKey(target.name) === sourceGroup);
                
                if (groupMatch) {
                    matchVal = groupMatch.index;
                } else {
                    const memMatch = rightOptions.find(opt => {
                        const saved = localStorage.getItem("map_" + opt.name);
                        if (!saved) return false;
                        const [sType, sVal] = saved.split(':');
                        return (row.type === 'source' && sType === 'SRC' && sVal === row.name) || 
                               (row.type === 'matrix' && sType === 'MAT' && sVal === row.name);
                    });
                    if (memMatch) matchVal = memMatch.index;
                }
            }
        }

        const isChecked = matchVal !== "-1" ? "checked" : "";
        const rowClass = matchVal !== "-1" ? "mapped-row" : "";
        let opts = `<option value="-1">-- Ignore / Select --</option>`;
        
        rightOptions.forEach(opt => {
            const sel = (opt.index === matchVal) ? "selected" : "";
            opts += `<option value="${opt.index}" ${sel}>${opt.name}</option>`;
        });
        
        let label = row.name + (row.type === 'matrix' ? ` <small style="color:#666">(${row.display})</small>` : "");
        const tr = document.createElement('tr');
        tr.className = rowClass;
        tr.setAttribute('data-search', row.name.toLowerCase());
        
        tr.innerHTML = `
            <td><b>${label}</b></td>
            <td style="text-align:center"><i class="fas fa-arrow-right" style="color:#9ca3af"></i></td>
            <td><select class="map-select" data-type="${row.type}" data-val="${row.val}" data-name="${row.name}" onchange="autoTick(this)">${opts}</select></td>
            <td style="text-align:center"><input type="checkbox" class="map-check" ${isChecked} onchange="updateMapStats()"></td>`;
        
        tbody.appendChild(tr);
    });
    
    updateMapOptionsVisibility();
    updateMapStats();
    return true; 
}

function updateMapOptionsVisibility() {
    const allSelects = document.querySelectorAll('.map-select');
    const usedValues = new Set();
    allSelects.forEach(sel => { if (sel.value !== "-1") usedValues.add(sel.value); });
    
    allSelects.forEach(sel => {
        sel.querySelectorAll('option').forEach(opt => {
            if (opt.value !== "-1") {
                opt.style.display = (usedValues.has(opt.value) && opt.value !== sel.value) ? "none" : "";
            }
        });
    });
}

function autoTick(selectEl) { 
    const checkbox = selectEl.parentElement.nextElementSibling.querySelector('input'); 
    
    if(selectEl.value !== "-1") {
        checkbox.checked = true;
        selectEl.closest('tr').classList.add('mapped-row');
        
        const targetName = selectEl.options[selectEl.selectedIndex].text;
        const type = selectEl.getAttribute('data-type') === 'source' ? 'SRC' : 'MAT';
        const name = selectEl.getAttribute('data-name');
        localStorage.setItem("map_" + targetName, `${type}:${name}`);
    } else {
        checkbox.checked = false;
        selectEl.closest('tr').classList.remove('mapped-row');
    }
    
    updateMapOptionsVisibility();
    updateMapStats(); 
}

function updateMapStats() {
    const rows = document.querySelectorAll('#mappingBody tr');
    let mappedCount = 0;
    let unmappedNames = [];

    rows.forEach(r => {
        const checkbox = r.querySelector('.map-check');
        const nameLabel = r.querySelector('td:first-child b'); 
        
        if(checkbox && checkbox.checked) {
            mappedCount++;
        } else if (nameLabel) {
            let cleanName = nameLabel.innerText.replace(/\s*\(.*?\)/, '').trim();
            unmappedNames.push(cleanName);
        }
    });

    const unmappedCount = rows.length - mappedCount;
    const alertEl = document.getElementById('unmappedList');
    
    if(alertEl) {
        if (unmappedCount === 0) {
            alertEl.style.backgroundColor = '#dcfce7'; 
            alertEl.style.color = '#166534';
            alertEl.style.borderColor = '#86efac';
            alertEl.innerHTML = `<i class="fas fa-check-circle"></i> All fields mapped`;
        } else {
            alertEl.style.backgroundColor = '#fee2e2'; 
            alertEl.style.color = '#b91c1c';
            alertEl.style.borderColor = '#fca5a5';
            
            let msg = "";
            if (unmappedNames.length <= 10) {
                msg = unmappedNames.join(", ");
            } else {
                msg = `${unmappedNames.slice(0, 10).join(", ")} + ${unmappedNames.length - 10} more`;
            }
            alertEl.innerHTML = `<i class="fas fa-exclamation-circle"></i> <strong>Unmapped:</strong> ${msg}`;
        }
    }
}

function resetMapping() {
    document.querySelectorAll('.map-select').forEach(s => { s.value = "-1"; });
    document.querySelectorAll('#mappingBody tr').forEach(r => { 
        r.classList.remove('mapped-row'); 
        r.querySelector('.map-check').checked = false; 
    });
    updateMapOptionsVisibility(); 
    updateMapStats(); 
}

function autoMapAgain() { renderMappingTable(); }

function filterMappingRows() {
    const filter = document.getElementById('mapSearchInput').value.toLowerCase();
    document.querySelectorAll('#mappingBody tr').forEach(r => {
        r.style.display = (r.getAttribute('data-search') || "").includes(filter) ? "" : "none";
    });
}

function clearSavedRules() {
    showModal("Clear Saved Rules?", "This will forget all auto-learned mappings.", "confirm", () => {
        const p = projects[activeProjectIdx];
        if(p.dataB) p.dataB.headers.forEach(h => localStorage.removeItem("map_" + h));
        resetMapping();
        showToast("Memory Cleared.");
    });
}

// ==========================================
// 11. ANALYSIS & DASHBOARD (Step 4)
// ==========================================

function generateResults() {
    // FIX: ADDED TRY-CATCH FOR ROBUSTNESS
    try {
        saveMappingFromUI(); 
        
        const p = projects[activeProjectIdx]; 
        
        if (!p.dataA || !p.dataB) {
            p.status = 'empty';
            p.step = 4;
            renderTopBar();
            renderDashboard();
            jumpToStep(4);
            return;
        }
        
        if (p.dataA.body.length !== p.dataB.body.length) {
             showModal("Row Mismatch", "Row counts must be equal.", 'error'); 
             return;
        }
        
        if (!p.mapping || p.mapping.length === 0) { 
            showModal("Error", "Please map at least one column.", 'error'); 
            return; 
        } 

        // DUPLICATE CHECK
        const seenTargets = new Set();
        for (const map of p.mapping) {
            if (seenTargets.has(map.idxB)) {
                showModal("Duplicate Mapping", `The field <b>"${map.name}"</b> is mapped more than once.<br>Please ensure each column is mapped only once.`, 'error');
                return; 
            }
            seenTargets.add(map.idxB);
        }
        
        p.status = 'done'; 
        p.step = 4; 
        
        renderTopBar(); 
        renderDashboard(); 
        jumpToStep(4); 
        
        const doTrim = document.getElementById('chkTrimResults')?.checked || false;
        showAnalysisReport(p, doTrim);
    } catch (err) {
        console.error("Analysis Failed:", err);
        showModal("Analysis Error", "An unexpected error occurred. Check console for details.", 'error');
    }
}

function saveMappingFromUI() { 
    const p = projects[activeProjectIdx]; 
    p.mapping = []; 
    
    // FIX: Iterate ROWS, not separate lists, to avoid index mismatch
    const rows = document.querySelectorAll('#mappingBody tr');
    
    rows.forEach(row => {
        const sel = row.querySelector('.map-select');
        const check = row.querySelector('.map-check');
        
        // Only map if both exist, checked, and value is valid
        if (check && check.checked && sel && sel.value !== "-1") {
            const targetIdx = parseInt(sel.value);
            const srcType = sel.getAttribute('data-type');
            const srcVal = sel.getAttribute('data-val'); 
            
            // Safety check for target header
            if (p.dataB.headers[targetIdx]) {
                p.mapping.push({ 
                    name: p.dataB.headers[targetIdx], 
                    idxB: targetIdx, 
                    targetType: srcType, 
                    targetVal: (srcType === 'source' ? parseInt(srcVal) : srcVal) 
                });
            }
        }
    });
}

function cleanAllSets() {
    let processed = 0;
    saveCurrentViewToProject(); 
    
    projects.forEach(p => {
        if (p.rawA && p.rawB) {
            p.dataA = parseExcelData(p.rawA, "1row"); 
            p.dataB = parseExcelData(p.rawB, "1row");
            if (p.dataA && p.dataB) {
                autoCleanData(p.dataA); 
                autoCleanData(p.dataB);
                p.status = 'ready'; 
                p.step = 2; 
                processed++;
            }
        }
    });
    
    if (processed === 0) {
        showModal("No Data Found", "No sets to clean.", 'error');
    } else { 
        showModal("Cleanup Complete", `Cleaned ${processed} sets.`, 'success'); 
        renderTopBar(); 
        loadProjectIntoView(activeProjectIdx); 
    }
}

function runAllComparisons() { 
    saveCurrentViewToProject(); 
    
    let processed = 0;
    let mismatchCount = 0;
    let missingDataLog = []; 

    const doTrim = document.getElementById('chkTrimResults')?.checked || false;
    const modeA = document.getElementById('headerModeA').value || "1row";
    const modeB = document.getElementById('headerModeB').value || "1row";

    projects.forEach(p => {
        if (p.rawA && !p.dataA) {
            p.dataA = parseExcelData(p.rawA, modeA);
            if(p.dataA) autoCleanData(p.dataA);
        }
        if (p.rawB && !p.dataB) {
            p.dataB = parseExcelData(p.rawB, modeB);
            if(p.dataB) autoCleanData(p.dataB);
        }

        const hasA = !!(p.dataA && p.dataA.body.length > 0);
        const hasB = !!(p.dataB && p.dataB.body.length > 0);

        if (hasA && hasB) {
            if (p.dataA.body.length === p.dataB.body.length) {
                autoMapProject(p);
                p.summary = calculateStats(p, doTrim);
                p.status = 'done'; 
                p.step = 4; 
                processed++;
            } else {
                p.status = 'ready'; 
                p.step = 2; 
                mismatchCount++;
            }
        } else {
            p.status = 'ready'; 
            p.step = 2; 
            
            if (!hasA && !hasB) {
                if (projects.length === 1) missingDataLog.push(`<strong>${p.name}</strong>: No data uploaded.`);
            } else if (!hasA) {
                missingDataLog.push(`<strong>${p.name}</strong>: Missing Form Data`);
            } else {
                missingDataLog.push(`<strong>${p.name}</strong>: Missing Sainpase/CPQ Data`);
            }
        }
    });

    renderTopBar(); 

    let title = "Analysis Complete";
    let type = "success";
    let msg = "";

    if (processed > 0) {
        msg += `<div style="margin-bottom:10px; color:#166534;"><i class="fas fa-check-circle"></i> Successfully processed <strong>${processed}</strong> set(s).</div>`;
    }

    if (mismatchCount > 0) {
        msg += `<div style="margin-bottom:10px; color:#b91c1c;"><i class="fas fa-exclamation-triangle"></i> <strong>${mismatchCount}</strong> set(s) skipped due to Row Mismatch.</div>`;
        type = "error"; 
    }

    if (missingDataLog.length > 0) {
        msg += `<div style="background:#fef2f2; border:1px solid #fecaca; padding:10px; border-radius:4px; color:#b91c1c; font-size:13px;">
                    <div style="font-weight:bold; margin-bottom:5px;">Missing Data Detected:</div>
                    <ul style="margin:0; padding-left:20px;">
                        ${missingDataLog.map(err => `<li>${err}</li>`).join('')}
                    </ul>
                </div>`;
        type = "error";
    }

    if (processed === 0 && mismatchCount === 0 && missingDataLog.length === 0) {
        showModal("No Data", "No valid data found to analyze.", 'error');
        return;
    }

    if (missingDataLog.length > 0 || projects.length > 1) {
        showOverview();
        setTimeout(() => showModal(processed > 0 ? "Analysis Report" : "Analysis Failed", msg, type), 300);
    } else {
        activeProjectIdx = 0;
        loadProjectIntoView(0);
        showToast("Analysis Complete");
    }
}

function autoMapProject(p) { 
    p.mapping = []; 
    const usedTargets = new Set(); 

    function getStandardKey(headerName) {
        let clean = headerName.toLowerCase()
            .replace(/\(?\boptional\b\)?/g, "")
            .replace(/[\u4e00-\u9fff]/g, "") 
            .replace(/[^a-z0-9]/g, "")      
            .trim();
        
        if (["footsize", "dimension", "dimensions", "dims", "measurement", "width", "height", "length", "wxh", "lxw"].some(k => clean.includes(k))) return "DIMENSION_GROUP";
        if (["quantity", "totalqty", "qty", "billshipquantity", "shippedqty", "invqty", "orderquantity", "orderqty", "units", "count"].includes(clean)) return "QTY_GROUP";
        if (["retailprice", "price", "retail", "retail1", "unitprice", "cost", "amount", "value"].includes(clean)) return "PRICE_GROUP";
        if (clean.includes("barcode") || ["upc", "gtin"].includes(clean)) return "BARCODE_GROUP";
        if (["productname", "itemname", "description", "desc", "shortdesc", "itemdesc"].includes(clean)) return "DESC_GROUP";
        if (["majorclass", "majclass", "class", "major", "category", "cat"].includes(clean)) return "CLASS_GROUP";
        if (["department", "dept", "division"].includes(clean)) return "DEPT_GROUP";
        if (["color", "colour", "clr", "colorname"].includes(clean)) return "COLOR_GROUP";
        if (["coo", "origin", "country", "countryoforigin"].includes(clean)) return "ORIGIN_GROUP";
        if (["po", "ponumber", "purchaseorder", "custpo"].includes(clean)) return "PO_GROUP";

        return clean; 
    }

    const rightOptions = p.dataB.headers.map((h, i) => ({ 
        name: h, 
        index: i,
        std: getStandardKey(h) 
    }));

    function tryMap(sourceName, sourceIndex, sourceType) {
        if (p.mapping.some(m => m.targetType === sourceType && m.targetVal === (sourceType === 'source' ? sourceIndex : sourceName))) return;

        const cleanKey = getStandardKey(sourceName);

        let match = rightOptions.find(opt => !usedTargets.has(opt.index) && opt.name.toLowerCase().trim() === sourceName.toLowerCase().trim());
        
        if (!match) {
             match = rightOptions.find(opt => !usedTargets.has(opt.index) && opt.std === cleanKey);
        }

        if (!match) {
            match = rightOptions.find(opt => !usedTargets.has(opt.index) && (
                (sourceName.toLowerCase().includes(opt.name.toLowerCase()) && opt.name.length > 3) || 
                (opt.name.toLowerCase().includes(sourceName.toLowerCase()) && sourceName.length > 3)
            ));
        }

        if (match) {
            p.mapping.push({ name: match.name, idxB: match.index, targetType: sourceType, targetVal: (sourceType === 'source' ? sourceIndex : sourceName) });
            usedTargets.add(match.index); 
            return true;
        }
        return false;
    }

    p.matrix.forEach(m => tryMap(m.key, null, 'matrix'));

    p.dataA.headers.forEach((h, i) => {
        let exactMatch = rightOptions.find(opt => !usedTargets.has(opt.index) && opt.name.toLowerCase().trim() === h.toLowerCase().trim());
        if (exactMatch) {
            p.mapping.push({ name: exactMatch.name, idxB: exactMatch.index, targetType: 'source', targetVal: i });
            usedTargets.add(exactMatch.index);
        }
    });

    p.dataA.headers.forEach((h, i) => {
        const isMapped = p.mapping.some(m => m.targetType === 'source' && m.targetVal === i);
        if (!isMapped) {
            tryMap(h, i, 'source');
        }
    });
}

// --- RENDER DASHBOARD ---
function renderDashboard() { 
    const p = projects[activeProjectIdx]; 
    const doTrim = document.getElementById('chkTrimResults')?.checked || false;
    
    // Safety Check: Handle Demo Mode or Missing Data
    if (!p.dataA || !p.dataB) {
        // Use diffCards (matching your HTML) instead of summaryCards
        const diffCards = document.getElementById('diffCards');
        if (diffCards) {
            diffCards.innerHTML = `<div class="field-card" style="border-left: 4px solid #ccc; width:100%"><div class="fc-head">Demo Mode</div><div class="fc-stats" style="color:#666">No Data Loaded Yet</div></div>`;
        }
        document.getElementById('globalStats').innerHTML = `<div class="big-stat"><div class="bs-val" style="color:#ccc">0</div><div class="bs-lbl">Rows</div></div>`;
        renderResultTables(0, doTrim);
        return;
    }

    const stats = calculateStats(p, doTrim);
    p.summary = stats;
    
    const maxRows = Math.max(p.dataA.body.length, p.dataB.body.length); 
    
    // Ensure we are targeting the correct element from your HTML file
    let diffCards = document.getElementById('diffCards');
    
    // Fallback: Create the container if it doesn't exist for some reason
    if (!diffCards) {
        diffCards = document.createElement('div');
        diffCards.id = 'diffCards';
        diffCards.className = 'cards-grid';
        // Insert it before the tables grid
        const tablesGrid = document.querySelector('.tables-grid');
        if (tablesGrid && tablesGrid.parentNode) {
            tablesGrid.parentNode.insertBefore(diffCards, tablesGrid);
        }
    }
    
    // Clear previous results
    diffCards.innerHTML = ""; 
    
    // Track if any mismatches are found to decide whether to show the "Success" message
    let hasMismatches = false;
    
    p.mapping.forEach(map => { 
        let match = 0, miss = 0; 
        const lowerName = map.name.toLowerCase();
        const isPrice = /price|cost|retail/i.test(lowerName);
        const isQty = /qty|quantity/i.test(lowerName); 
        
        for (let i = 0; i < maxRows; i++) { 
            let vB = (p.dataB.body[i]?.[map.idxB] || "").toString(); 
            let vA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").toString(); 
            
            if(doTrim) { vB = vB.replace(/\s+/g, ''); vA = vA.replace(/\s+/g, ''); } 
            else { vB = vB.trim(); vA = vA.trim(); }

            let normA = vA.toLowerCase();
            let normB = vB.toLowerCase(); 
            let equal = false; 
            
            if (isPrice) equal = (normA.replace(/[$,]/g,'') === normB.replace(/[$,]/g,'')); 
            else if (isQty) equal = (normA.replace(/[\,]/g,'') === normB.replace(/[\,]/g,'')); 
            else equal = (normA === normB); 
            
            if (equal) match++; else miss++; 
        } 
        
        // --- LOGIC CHANGE: Only render the card if mismatches > 0 ---
        if (miss > 0) {
            hasMismatches = true;
            // Since we only show errors, we use the warning color
            const cls = 'bg-warn'; 
            diffCards.innerHTML += `<div class="field-card ${cls}"><div class="fc-head">${map.name}</div><div class="fc-stats"><span style="color:#10b981">✓ ${match}</span><span style="color:#ef4444">✗ ${miss}</span></div></div>`; 
        }
    }); 
    
    // --- LOGIC CHANGE: Show Success Message if no mismatches were added ---
    if (!hasMismatches) {
        diffCards.innerHTML = `
            <div style="grid-column: 1 / -1; text-align: center; padding: 40px; background: white; border-radius: 12px; border: 1px solid #bbf7d0; display:flex; flex-direction:column; align-items:center;">
                <i class="fas fa-check-circle" style="font-size: 48px; color: #22c55e; margin-bottom: 15px;"></i>
                <h3 style="margin: 0; color: #15803d; font-size: 20px;">All Fields Matched!</h3>
                <p style="margin: 10px 0 0 0; color: #64748b;">No mismatches found in the mapped columns.</p>
            </div>`;
    }
    
    document.getElementById('globalStats').innerHTML = `
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${maxRows}</div><div class="bs-lbl">Rows</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${p.mapping.length}</div><div class="bs-lbl">Fields</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#10b981">${stats.matches}</div><div class="bs-lbl">Matches</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#ef4444">${stats.mismatches}</div><div class="bs-lbl">Mismatches</div></div>`; 
    
    renderResultTables(maxRows, doTrim); 
}

function calculateStats(p, doTrim) {
    if (!p.dataA || !p.dataB || !p.mapping) return { matches: 0, mismatches: 0 };
    let matches = 0, mismatches = 0;
    const rows = Math.max(p.dataA.body.length, p.dataB.body.length);
    
    p.mapping.forEach(map => {
        const isPrice = /price|cost|retail/i.test(map.name.toLowerCase());
        const isQty = /qty|quantity/i.test(map.name.toLowerCase());
        
        for(let i=0; i<rows; i++) {
            let vB = (p.dataB.body[i]?.[map.idxB] || "").toString();
            let vA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").toString();
            
            if(doTrim) { vB = vB.replace(/\s+/g, ''); vA = vA.replace(/\s+/g, ''); } 
            else { vB = vB.trim(); vA = vA.trim(); }
            
            let equal = false;
            if (isPrice) equal = (vA.toLowerCase().replace(/[$,]/g,'') === vB.toLowerCase().replace(/[$,]/g,''));
            else if (isQty) equal = (vA.toLowerCase().replace(/[\,\s]/g,'') === vB.toLowerCase().replace(/[\,\s]/g,''));
            else equal = (vA.toLowerCase() === vB.toLowerCase());
            
            if(equal) matches++; else mismatches++;
        }
    });
    return { matches, mismatches };
}

function showAnalysisReport(p, doTrim) {
    const stats = calculateStats(p, doTrim);
    
    if (stats.mismatches === 0) {
        const successHtml = `
            <div style="text-align:center;">
                <p style="font-size: 15px; margin-bottom: 20px; color: #374151;">
                    All <strong>${p.dataA.body.length}</strong> rows match perfectly across all columns.
                </p>
                <div style="background-color: #ecfdf5; color: #047857; padding: 15px; border-radius: 6px; font-weight: bold; border: 1px solid #a7f3d0; letter-spacing: 0.5px;">
                    NO MISMATCHES FOUND
                </div>
            </div>
        `;
        showModal("Analysis Complete: Perfect Match!", successHtml, 'success');
        return;
    }

    let tableRows = "";
    
    p.mapping.forEach(map => {
        let colErrorCount = 0;
        const totalRows = p.dataA.body.length;

        for (let i = 0; i < totalRows; i++) {
            let valA = map.targetType === 'matrix' 
                ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") 
                : (p.dataA.body[i]?.[map.targetVal] || "").toString();

            let valB = (p.dataB.body[i]?.[map.idxB] || "").toString();

            if (doTrim) {
                valA = valA.replace(/\s+/g, '');
                valB = valB.replace(/\s+/g, '');
            } else {
                valA = valA.trim();
                valB = valB.trim();
            }

            const lowerName = map.name.toLowerCase();
            const isPrice = /price|cost|retail/i.test(lowerName);
            const isQty = /qty|quantity/i.test(lowerName);

            let equal = false;
            if (isPrice) {
                equal = (valA.toLowerCase().replace(/[$,]/g, '') === valB.toLowerCase().replace(/[$,]/g, ''));
            } else if (isQty) {
                equal = (valA.toLowerCase().replace(/[\,\s]/g, '') === valB.toLowerCase().replace(/[\,\s]/g, ''));
            } else {
                equal = (valA.toLowerCase() === valB.toLowerCase());
            }

            if (!equal) colErrorCount++;
        }

        if (colErrorCount > 0) {
            tableRows += `
                <tr>
                    <td style="padding:10px; border-bottom:1px solid #f3f4f6; color:#374151;">${map.name}</td>
                    <td style="padding:10px; border-bottom:1px solid #f3f4f6; text-align:right; font-weight:bold; color:#ef4444;">${colErrorCount}</td>
                </tr>
            `;
        }
    });

    const reportHtml = `
        <div style="text-align:left;">
            <p style="margin-bottom: 15px; font-size:15px; color:#374151;">
                Found <strong>${stats.mismatches}</strong> total mismatches in <strong>${p.dataA.body.length}</strong> rows.
            </p>
            
            <div style="max-height: 300px; overflow-y:auto; border:1px solid #e5e7eb; border-radius:6px;">
                <table style="width:100%; border-collapse:collapse; font-size:14px;">
                    <thead style="position:sticky; top:0; background:#f9fafb;">
                        <tr style="border-bottom:1px solid #e5e7eb; color:#b91c1c;">
                            <th style="text-align:left; padding:10px; font-weight:600;">Column Name</th>
                            <th style="text-align:right; padding:10px; font-weight:600;">Mismatches</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    showModal("Analysis Complete: Mismatches Found", reportHtml, 'error');
}

function toggleMismatchView() {
    const onlyErrors = document.getElementById('chkShowErrors').checked;
    const styleId = 'mismatch-style';
    let styleTag = document.getElementById(styleId);
    
    if (onlyErrors) {
        if (!styleTag) {
            styleTag = document.createElement('style');
            styleTag.id = styleId;
            styleTag.innerHTML = `.perfect-row { display: none !important; }`; 
            document.head.appendChild(styleTag);
        }
    } else {
        if (styleTag) styleTag.remove();
    }
}

function renderResultTables(maxRows, doTrim) { 
    const p = projects[activeProjectIdx]; 
    const tA = document.getElementById('renderTableA');
    const tB = document.getElementById('renderTableB'); 
    
    if (!p.dataA || !p.dataB) {
        if(tA) tA.innerHTML = "<thead><tr><th>Form Data</th></tr></thead><tbody><tr><td style='color:#ccc; padding:20px; text-align:center;'>No Data Available</td></tr></tbody>";
        if(tB) tB.innerHTML = "<thead><tr><th>Sainpase Data</th></tr></thead><tbody><tr><td style='color:#ccc; padding:20px; text-align:center;'>No Data Available</td></tr></tbody>";
        return;
    }

    let hA = "<thead><tr><th>#</th>" + p.dataA.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    let hB = "<thead><tr><th>#</th>" + p.dataB.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    
    let bA = ""; 
    let bB = ""; 
    
    const mapLookup = {}; 
    p.mapping.forEach(m => mapLookup[m.idxB] = m); 
    
    const reverseLookup = {}; 
    p.mapping.forEach(m => { 
        if (m.targetType === 'source') { 
            if (!reverseLookup[m.targetVal]) reverseLookup[m.targetVal] = []; 
            reverseLookup[m.targetVal].push(m); 
        } 
    }); 
    
    function checkEqual(valA, valB, fieldName) {
        let vA = doTrim ? valA.replace(/\s+/g, '') : valA.trim();
        let vB = doTrim ? valB.replace(/\s+/g, '') : valB.trim();
        let normA = vA.toLowerCase();
        let normB = vB.toLowerCase(); 
        
        const lowerName = fieldName.toLowerCase();
        const isPrice = /price|cost|retail/i.test(lowerName);
        const isQty = /qty|quantity/i.test(lowerName); 
        
        if (isPrice) {
            return (normA.replace(/[$,]/g,'') === normB.replace(/[$,]/g,'')); 
        } else if (isQty) {
            return (normA.replace(/[\,]/g,'') === normB.replace(/[\,]/g,'')); 
        } else {
            return (normA === normB); 
        }
    }

    for (let i = 0; i < maxRows; i++) { 
        let rA = `<td>${i+1}</td>`; 
        
        for (let cA = 0; cA < p.dataA.headers.length; cA++) { 
            let rawA = (p.dataA.body[i]?.[cA] || "").toString();
            let displayA = rawA;
            let cls = ""; 
            
            if (reverseLookup[cA]) { 
                let allMatch = true;
                let comparedAgainst = ""; 
                
                reverseLookup[cA].forEach(map => { 
                    let rawB = (p.dataB.body[i]?.[map.idxB] || "").toString(); 
                    comparedAgainst = rawB; 
                    
                    if (!checkEqual(rawA, rawB, map.name)) {
                        allMatch = false; 
                    }
                }); 
                
                cls = allMatch ? "match" : "diff"; 
                if (!allMatch) {
                    displayA = getVisualDiff(rawA, comparedAgainst);
                }
            } 
            rA += `<td class="${cls}">${displayA}</td>`; 
        } 
        
        let rB = `<td>${i+1}</td>`; 
        p.dataB.headers.forEach((_, colIdx) => { 
            let rawB = (p.dataB.body[i]?.[colIdx] || "").toString();
            let displayB = rawB;
            let cls = ""; 
            
            if (mapLookup.hasOwnProperty(colIdx)) { 
                let map = mapLookup[colIdx];
                let rawA = map.targetType === 'matrix' 
                    ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") 
                    : (p.dataA.body[i]?.[map.targetVal] || "").toString(); 
                
                if (checkEqual(rawA, rawB, map.name)) {
                    cls = "match";
                } else {
                    cls = "diff";
                    displayB = getVisualDiff(rawB, rawA);
                }
            } 
            rB += `<td class="${cls}">${displayB}</td>`; 
        }); 

        const isRowError = rA.includes('class="diff"') || rB.includes('class="diff"');
        const rowClass = isRowError ? "issue-row" : "perfect-row";
        
        bA += `<tr class="${rowClass}">${rA}</tr>`; 
        bB += `<tr id="rowB-${i}" class="${rowClass}">${rB}</tr>`; 
    } 
    
    tA.innerHTML = hA + bA + "</tbody>"; 
    tB.innerHTML = hB + bB + "</tbody>"; 
}

// ==========================================
// 12. HELPER UTILITIES
// ==========================================

function findHeaderRowIndex(data) {
    const qtyRegex = /qty|quantity/i;
    for (let i = 0; i < Math.min(data.length, 60); i++) {
        if (data[i] && data[i].some(cell => cell && qtyRegex.test(String(cell).trim()))) {
            return i;
        }
    }
    return -1;
}



function getVisualDiff(mainText, compareText) {
    if (!mainText) return ""; 
    if (!compareText) return mainText; 
    
    const wordsA = mainText.toString().split(/\s+/);
    const wordsB = compareText.toString().split(/\s+/);
    let html = "";
    
    wordsA.forEach((word, i) => {
        const otherWord = wordsB[i] || "";
        if (word.toLowerCase() !== otherWord.toLowerCase()) {
            html += `<span class="diff-word">${word}</span> `;
        } else {
            html += word + " ";
        }
    });
    return html.trim();
}

function autoCleanData(data) {
    if (!data || !data.body) return;

    const keepIndices = data.headers
        .map((h, i) => {
            const val = h ? String(h).trim() : "";
            if (val === "" || val === "#") return -1;
            return i;
        })
        .filter(i => i !== -1);

    if (keepIndices.length < data.headers.length) {
        data.headers = keepIndices.map(i => data.headers[i]);
        data.body = data.body.map(row => keepIndices.map(i => row[i]));
    }

    const qtyRegex = /qty|quantity|shipped|billed|units|pcs/i;
    const qtyIndices = [];
    
    data.headers.forEach((h, i) => { 
        if (h && qtyRegex.test(h.toString().toLowerCase())) qtyIndices.push(i); 
    });
    
    if (qtyIndices.length === 0) return;
    
    data.body = data.body.filter(row => {
        if (!row.some(cell => cell && cell.toString().trim() !== "")) return false;
        
        if (row.join(" ").toLowerCase().includes("upc") && row.join(" ").toLowerCase().includes("style")) return false; 
        
        for (let idx of qtyIndices) {
            let numVal = parseFloat(String(row[idx] || "").toLowerCase().replace(/[, \s]/g, '').replace(/pcs/g, ''));
            if (!isNaN(numVal) && numVal > 0) return true;
        }
        return false; 
    });
}

function parseExcelData(raw, mode) { 
    if (!raw || !raw.trim()) return null; 
    
    let text = raw.trim().replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let rows = [], currentRow = [], currentCell = "", insideQuote = false;
    
    for (let i = 0; i < text.length; i++) {
        let char = text[i], nextChar = text[i+1];
        if (insideQuote) {
            if (char === '"') { 
                if (nextChar === '"') { currentCell += '"'; i++; } 
                else { insideQuote = false; } 
            } else { 
                currentCell += char; 
            }
        } else {
            if (char === '"') { 
                if (currentCell.length === 0) insideQuote = true; 
                else currentCell += char; 
            } 
            else if (char === '\t') { 
                currentRow.push(currentCell.trim()); currentCell = ""; 
            } 
            else if (char === '\n') { 
                currentRow.push(currentCell.trim()); rows.push(currentRow); currentRow = []; currentCell = ""; 
            } 
            else { 
                currentCell += char; 
            }
        }
    }
    if (currentCell || currentRow.length > 0) { 
        currentRow.push(currentCell.trim()); 
        rows.push(currentRow); 
    }
    if (rows.length < 1) return null;

    let ignoreFirstColumn = (rows[0][0] && rows[0][0].toLowerCase().includes("sku information"));
    if (ignoreFirstColumn) rows.shift(); 
    
    let headers = [], bodyStartIndex = 1; 
    
    if (mode === "2rows" && rows.length >= 2) { 
        const maxCols = Math.max(rows[0].length, rows[1].length); 
        for (let i = 0; i < maxCols; i++) {
            headers.push(cleanHeader(rows[1][i] || "") || cleanHeader(rows[0][i] || "")); 
        }
        bodyStartIndex = 2; 
    } else { 
        headers = rows[0].map(h => cleanHeader(h)); 
    } 
    
    let bodyRaw = rows.slice(bodyStartIndex);
    let body = bodyRaw.map(r => { 
        while (r.length < headers.length) r.push(""); 
        return r; 
    }); 
    
    if (ignoreFirstColumn) { 
        headers.shift(); 
        body = body.map(r => r.slice(1)); 
    } 
    return { headers, body }; 
}

function cleanHeader(text) { 
    if (!text) return "";
    
    let clean = text.trim()
        .replace(/^"|"$/g, '') 
        .replace(/\s*max\s+\d+\s*(digits|chars|characters|digit)/gi, "")
        .replace(/_/g, " ")
        .replace(/[\r\n]+/g, " ")
        .replace(/\s*\([^\)]*\)/g, "") 
        .replace(/\*/g, "")
        .trim(); 
    
    if(clean.includes("SIZEStyle")) return "Style"; 
    
    return clean;
}



function showToast(msg) {
    const t = document.getElementById('toast'); 
    t.querySelector('span').innerHTML = msg || "Done";
    t.classList.remove('hidden'); 
    setTimeout(() => { t.classList.add('hidden'); }, 3000); 
}

// --- DRAG AND DROP LOGIC ---

document.addEventListener('DOMContentLoaded', () => {
    setupDragDrop('cardSource', 'A');
    setupDragDrop('cardTarget', 'B'); 
});

function setupDragDrop(elementId, side) {
    const card = document.getElementById(elementId);
    if (!card) return;

    card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        card.classList.add('drag-active');
    });

    card.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        card.classList.remove('drag-active');
    });

    card.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        card.classList.remove('drag-active');

        const files = e.dataTransfer.files;
        if (files && files.length > 0) {
            const fakeInput = { files: files, value: "" };

            if (side === 'A') {
                handleBulkUpload(fakeInput);
            } else {
                handleCPQUpload(fakeInput);
            }
        }
    });
}
