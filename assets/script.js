// ==========================================
// 1. GLOBAL VARIABLES & INITIALIZATION
// ==========================================

let projects = [];
let activeProjectIdx = 0;
let isOverviewMode = false;
const MAX_TOTAL_SETS = 50;
let manualWorkbook = null;
let manualFilename = "";
let currentSheetName = "";
let selectedHeaderRowIndex = -1;

// Selection State
let isDragging = false;
let selectionStartRow = -1;
let selectionEndRow = -1; 

// Undo History Variables
let deletedItem = null;      
let deletedItemIdx = -1;     
let deletedRowsHistory = []; 


// NEW: Memory for the last uploaded file
let lastUploadedWorkbook = null;
let lastUploadedFilename = "";
let editingProjectIndex = -1;

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

    // --- FIX: Hide Restore Button ---
    const restoreBtn = document.getElementById('btnRestoreSet');
    if (restoreBtn) {
        restoreBtn.style.display = 'none';
        restoreBtn.innerHTML = ''; // Clear text content too just in case
    }
    
    // --- FIX: Clear Custom Keyword Input (Added in previous step) ---
    const customKeyInput = document.getElementById('customQtyKey');
    if (customKeyInput) customKeyInput.value = "";

    // 3. Reset UI State (Sidebar & Tabs)
    const sidebar = document.querySelector('.sidebar');
    if (sidebar) sidebar.classList.remove('overview-locked');
    
    const tabOverview = document.getElementById('tabOverview');
    if (tabOverview) tabOverview.classList.remove('active');
    
    const overviewSection = document.getElementById('overviewSection');
    if (overviewSection) overviewSection.style.display = 'none';

    // 4. Reset to Step 1 visually
    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    const step1 = document.getElementById('step1');
    if (step1) step1.style.display = 'block';
    
    // Reset Sidebar Navigation Styles
    document.querySelectorAll('.v-step').forEach(el => el.classList.remove('active'));
    const navStep1 = document.getElementById('navStep1');
    if (navStep1) navStep1.classList.add('active');

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


function createSet(customName = null, workbook = null, sheetName = null, fileName = "") {
    const id = projects.length + 1;
    projects.push({
        name: customName || `Set ${id}`,
        sourceWorkbook: workbook, 
        originalSheetName: sheetName,
        fileName: fileName || "",
        status: 'empty', 
        rawA: "", rawB: "", rawMatrix: "",
        dataA: null, dataB: null, matrix: [],
        mapping: [], step: 1, showMatrix: false,
        summary: { matches: 0, mismatches: 0 },
        // --- ENHANCEMENT: INDEPENDENT SETTINGS MEMORY ---
        settings: {
            modeA: '1row',
            modeB: '1row',
            trimResults: false,
            showErrors: false
        }
    });
}

// ==========================================
// 3. FILE UPLOAD HANDLERS
// ==========================================

// ==========================================
// 3. FILE UPLOAD HANDLERS
// ==========================================

function handleBulkUpload(input) {
    if (!input.files || input.files.length === 0) return;

    // Save current work before processing
    saveCurrentViewToProject(); 
    
    // Safety: Clear projects if it's the very first empty load
    const currentA = document.getElementById('tableA').value.trim();
    const currentB = document.getElementById('tableB').value.trim(); 
    if (projects.length === 1 && projects[0].status === 'empty' && !projects[0].rawA && !currentA && !currentB) {
        projects = []; 
        activeProjectIdx = -1;
    }

    // --- RESTORED: GET CUSTOM KEYWORD ---
    const customKeyInput = document.getElementById('customQtyKey');
    const customKey = customKeyInput ? customKeyInput.value.trim() : "";

    // --- ENHANCEMENT: SORT MULTIPLE FILES BY DOWNLOAD TIME (Oldest to Newest) ---
    const files = Array.from(input.files).sort((a, b) => a.lastModified - b.lastModified);
    
    let totalCreated = 0;
    
    // Memory for manual fallback
    let tempWorkbook = null;
    let tempFilename = "";

    const processNextFile = (fileIdx) => {
        if (fileIdx >= files.length) {
            input.value = ""; // Reset input
            
            // Save last file for manual fallback
            if (tempWorkbook) {
                lastUploadedWorkbook = tempWorkbook;
                lastUploadedFilename = tempFilename;
            }

            // --- RESULT HANDLING ---
            renderTopBar();
            if (totalCreated > 0) {
                // Switch to the FIRST newly added set
                const firstNewSetIdx = projects.length - totalCreated;
                switchProject(firstNewSetIdx, true); 

                const msg = `✅ Successfully loaded <strong>${totalCreated}</strong> sheets from <strong>${files.length}</strong> file(s).`;
                showModal("Import Complete", msg, "success");
            } else {
                if (projects.length === 0) { createSet(); activeProjectIdx = 0; }
                
                // RESTORED: Trigger manual mapping fallback if auto fails
                if (typeof openManualMapper === "function") {
                    openManualMapper(lastUploadedWorkbook, lastUploadedFilename);
                }
            }
            return;
        }

        const file = files[fileIdx];
        tempFilename = file.name;
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                tempWorkbook = workbook;
                
                // --- RESTORED: SHEET SORTING LOGIC (A-Z inside the file) ---
                let sheetNamesToProcess = [...workbook.SheetNames];
                const sortToggle = document.getElementById('sortSheetsToggle');
                
                if (sortToggle && sortToggle.checked) {
                    sheetNamesToProcess.sort((a, b) => a.localeCompare(b, undefined, {numeric: true}));
                }
                
                sheetNamesToProcess.forEach((sheetName) => {
                    // Skip hidden sheets
                    if (workbook.Workbook && workbook.Workbook.Sheets) {
                        const sMeta = workbook.Workbook.Sheets.find(s => s.name === sheetName);
                        if (sMeta && (sMeta.Hidden !== 0 || sMeta.state === 'hidden')) return;
                    }

                    const sheet = workbook.Sheets[sheetName];
                    
                    // --- RESTORED: SMART BARCODE EXTRACTOR ---
                    let rawData = typeof extractSmartExcelData === "function" 
                        ? extractSmartExcelData(sheet) 
                        : XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false }); 
                    // --- NEW: HIDDEN ROW REMOVAL ---
                    if (sheet['!rows'] && rawData.length > 0) {
                        rawData = rawData.filter((row, rowIndex) => {
                            const rMeta = sheet['!rows'][rowIndex];
                            // Filter out if row is explicitly hidden, or if height is 0
                            const isHidden = rMeta && (rMeta.hidden === true || rMeta.hidden === 1 || rMeta.hpx === 0 || rMeta.ht === 0);
                            return !isHidden;
                        });
                    }
                    
                    // --- RESTORED: ADVANCED HIDDEN COLUMN REMOVAL ---
                    if (sheet['!cols'] && rawData.length > 0) {
                        const hiddenIndices = new Set();
                        let startIdx = -1;
                        let endIdx = -1;
                        
                        for(let r=0; r<Math.min(rawData.length, 30); r++) {
                            const row = rawData[r];
                            if(!row || !Array.isArray(row)) continue;
                            
                            if (startIdx === -1) {
                                startIdx = row.findIndex(c => {
                                    const s = String(c).toUpperCase();
                                    return s.includes("FEATURE FIELD 1") || s.includes("FEATURE ICON 1") || (s.includes("FEATURE CALL") && s.includes("1"));
                                });
                            }
                            if (endIdx === -1) {
                                 endIdx = row.findIndex(c => {
                                    const s = String(c).toUpperCase();
                                    return s.includes("TOTAL QTY") || s.includes("TOTAL QUANTITY") || (s.includes("TOTAL") && s.includes("QTY"));
                                 });
                            }
                            if (startIdx !== -1 && endIdx !== -1) break;
                        }

                        sheet['!cols'].forEach((col, i) => {
                            if (!col) return;
                            const isExplicitHidden = col.hidden === true || col.hidden === 1;
                            const isZeroWidth = (col.wpx != null && col.wpx < 1) || (col.width != null && col.width < 0.1);
                            
                            if (isExplicitHidden || isZeroWidth) {
                                let keep = false;
                                if (startIdx !== -1 && endIdx !== -1) {
                                    if (i >= startIdx && i <= endIdx) keep = true;
                                }
                                if (!keep && startIdx !== -1 && (i > startIdx && i < startIdx + 6)) {
                                    keep = true;
                                }
                                if (!keep) {
                                    hiddenIndices.add(i);
                                }
                            }
                        });

                        if (hiddenIndices.size > 0) {
                            rawData = rawData.map(row => row.filter((_, i) => !hiddenIndices.has(i)));
                        }
                    }

                    // --- RESTORED: CUSTOM KEYWORD DATA EXTRACTION ---
                    const cleanTable = extractTableStrict(rawData, customKey); 

                    // RESTORED: BLANK QUANTITY CHECK
                    let hasValidData = false;
                    if (cleanTable.length > 1) {
                        const headerRow = cleanTable[0];
                        // ⚡ FIX: Updated validation regex to match the ENTIRE Qty Group suite ⚡
                        const searchRegex = customKey ? new RegExp(customKey, "i") : /qty|quantity|total\s*qty|bill\s*(?:&|and)?\s*ship\s*(?:qty|quantity)|round\s*up|order\s*(?:qty|quantity)/i;
                        const qtyColIdx = headerRow.findIndex(h => searchRegex.test(String(h)));

                        if (qtyColIdx !== -1) {
                            for (let r = 1; r < cleanTable.length; r++) {
                                const val = cleanTable[r][qtyColIdx];
                                if (val && String(val).trim() !== "" && String(val).trim() !== "0") {
                                    hasValidData = true;
                                    break;
                                }
                            }
                        } else {
                            hasValidData = true; 
                        }
                    }

                    if (hasValidData) {
                        totalCreated++;
                        
                        // --- ENHANCEMENT: PURE SHEET NAMING LOGIC ---
                        let finalName = sheetName; 
                        let counter = 1;
                        
                        // Safety check: Only rename if the EXACT SAME FILE has two sheets with the same name
                        while(projects.some(p => p.name === finalName && p.fileName === file.name)) {
                            finalName = `${sheetName} (${counter})`;
                            counter++;
                        }

                        // Create independent tab
                        createSet(finalName, null, sheetName, file.name); 
                        let p = projects[projects.length - 1];
                        // ---------------------------------------------
                        
                        p.rawA = arrayToTSV(cleanTable);
                        
                        // Matrix extraction
                        try {
                            const headerIdx = findHeaderRowIndex(rawData);
                            if(headerIdx > 0) {
                                const matrixData = extractMatrixData(rawData, headerIdx);
                                if (matrixData.length > 0) {
                                    p.rawMatrix = matrixData.map(m => `${m.key}: ${m.val}`).join("\n");
                                    p.matrix = parseMatrixString(p.rawMatrix);
                                    p.showMatrix = true;
                                }
                            }
                        } catch(e) {} 

                        p.status = 'ready';
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

    processNextFile(0); // Kickoff the loop
}


function retryWithKeyword() {
    const customKeyInput = document.getElementById('customQtyKey');
    const customKey = customKeyInput ? customKeyInput.value.trim() : "";

    if (!customKey) {
        showToast("Please enter a keyword");
        return;
    }

    if (!lastUploadedWorkbook) {
        showModal("No File", "Please upload a file first.", "error");
        return;
    }

    const workbook = lastUploadedWorkbook;
    let matchedSheets = [];

    showToast(`Scanning for "${customKey}"...`);

    // 1. Scan ALL visible sheets
    workbook.SheetNames.forEach((sheetName) => {
        // Skip hidden sheets
        if (workbook.Workbook && workbook.Workbook.Sheets) {
            const sMeta = workbook.Workbook.Sheets.find(s => s.name === sheetName);
            if (sMeta && (sMeta.Hidden !== 0)) return;
        }

        const sheet = workbook.Sheets[sheetName];
        let rawData = extractSmartExcelData(sheet);

        // Check if this sheet has the keyword
        const cleanTable = extractTableStrict(rawData, customKey);
        
        // If valid table found (more than just header)
        if (cleanTable.length > 1) {
            matchedSheets.push({ name: sheetName, data: cleanTable });
        }
    });

    if (matchedSheets.length === 0) {
        showModal("Not Found", `No sheets found containing column: <strong>${customKey}</strong>`, "error");
    } else {
        // 2. Trigger the Selector Modal
        showSheetSelector(matchedSheets, customKey);
    }
}

function showSheetSelector(matches, keyword) {
    // We will reuse the 'manualSelectModal' for this purpose
    const modal = document.getElementById('manualSelectModal');
    const tableArea = document.getElementById('manualRawTable'); 
    const titleArea = modal.querySelector('h3');
    const instruction = document.getElementById('manualInstruction');
    const confirmBtn = document.getElementById('btnManualConfirm');
    
    // 1. Configure Modal UI
    modal.style.display = 'flex';
    titleArea.innerHTML = `<i class="fas fa-search"></i> Found "${keyword}" in ${matches.length} Sheets`;
    instruction.innerHTML = `<span style="color:#2563eb">Select sheets to import as separate sets:</span>`;
    
    // 2. Build the Selection List
    let html = `
    <div style="padding:20px; height:100%; overflow-y:auto;">
        <div style="margin-bottom:15px; padding-bottom:10px; border-bottom:1px solid #e2e8f0;">
            <label style="font-weight:bold; cursor:pointer; display:flex; align-items:center; gap:10px;">
                <input type="checkbox" id="chkSelectAll" onchange="toggleAllSheets(this)"> 
                <span>Select All Sheets</span>
            </label>
        </div>
        <div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap:12px;">`;

    matches.forEach((m, idx) => {
        html += `
        <label style="background:#f8fafc; border:1px solid #cbd5e1; padding:12px; border-radius:6px; cursor:pointer; display:flex; align-items:center; gap:10px; transition:all 0.2s;">
            <input type="checkbox" class="sheet-chk" value="${idx}" checked>
            <div>
                <div style="font-weight:600; color:#334155;">${m.name}</div>
                <div style="font-size:11px; color:#64748b;">${m.data.length} rows found</div>
            </div>
        </label>`;
    });
    html += `</div></div>`;
    
    // Inject HTML (replacing the table)
    const container = tableArea.parentElement;
    // Hide the original table container temporarily
    tableArea.style.display = 'none';
    
    // Check if we already appended a custom list div, if so remove it to refresh
    const existingList = document.getElementById('customSheetList');
    if (existingList) existingList.remove();

    const listDiv = document.createElement('div');
    listDiv.id = 'customSheetList';
    listDiv.style.flex = "1";
    listDiv.style.overflow = "hidden";
    listDiv.innerHTML = html;
    container.appendChild(listDiv);

    // 3. Configure Confirm Button
    confirmBtn.disabled = false;
    confirmBtn.innerHTML = "Import Selected";
    confirmBtn.style.background = "#2563eb";
    confirmBtn.style.cursor = "pointer";

    // helper for Select All
    window.toggleAllSheets = function(source) {
        document.querySelectorAll('.sheet-chk').forEach(c => c.checked = source.checked);
    };

    // 4. Handle Import Action
    confirmBtn.onclick = function() {
        const checkboxes = document.querySelectorAll('.sheet-chk:checked');
        if (checkboxes.length === 0) {
            alert("Please select at least one sheet.");
            return;
        }

        // Import Loop
        checkboxes.forEach((chk, i) => {
            const match = matches[parseInt(chk.value)];
            
            // Use current project for the first one, create new projects for others
            let targetIdx = activeProjectIdx;
            
            if (i > 0) {
                // Add a new set
                projects.push({
                    id: Date.now() + i,
                    name: `Set ${projects.length + 1}`,
                    step: 1,
                    status: 'empty',
                    rawA: "", rawB: "",
                    cleanA: [], cleanB: [],
                    mapping: [],
                    results: null
                });
                targetIdx = projects.length - 1;
            }

            const p = projects[targetIdx];
            p.name = match.name;
            p.fileName = lastUploadedFilename;
            p.rawA = arrayToTSV(match.data);
            p.status = 'ready'; // Mark as ready immediately
            p.step = 1;
        });

        // Cleanup
        document.getElementById('customSheetList').remove();
        tableArea.style.display = 'block'; // Restore original table view for manual select
        closeManualModal();
        
        // Refresh View
        loadProjectIntoView(activeProjectIdx);
        renderTopBar();
        showToast(`Imported ${checkboxes.length} sheets successfully!`);

        // Restore button for Manual Select
        confirmBtn.onclick = confirmManualImport;
        confirmBtn.innerHTML = "Confirm & Import";
    };
}

// --- SMART EXCEL EXTRACTOR ---
// Fixes Barcode scientific notation without ruining formatted Quantities!
function extractSmartExcelData(sheet) {
    // 1. Get visual formatted text (Good for QTY, Prices, Dates)
    let dataVisual = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    
    // 2. Get exact pure numbers (Good for Barcodes)
    let dataRaw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    
    // 3. Merge them: If visual text is ruined by scientific notation, use the raw number!
    for (let r = 0; r < dataVisual.length; r++) {
        if (!dataVisual[r] || !dataRaw[r]) continue;
        for (let c = 0; c < dataVisual[r].length; c++) {
            let textVal = String(dataVisual[r][c]).trim();
            
            // If it looks like "8.84969E+11" (Scientific Notation)
            if (/^\d+(\.\d+)?[Ee]\+\d+$/.test(textVal)) {
                dataVisual[r][c] = dataRaw[r][c]; // Pull the exact unformatted number
            }
        }
    }
    return dataVisual;
}


// ==========================================
// 4. DATA EXTRACTION LOGIC
// ==========================================

function extractTableStrict(data, customKeyword = null) {
    // 1. Safety Check
    if (!data || !Array.isArray(data) || data.length === 0) return [];
    data = data.filter(r => Array.isArray(r));

    let startIndex = -1;
    
    // ⚡ FIX: Comprehensive QTY_GROUP Regex for identifying tables ⚡
    let qtyRegex;
    if (customKeyword) {
        const escaped = customKeyword.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const flexSpace = escaped.replace(/\s+/g, '\\s*'); 
        qtyRegex = new RegExp(flexSpace, "i");
    } else {
        qtyRegex = /qty|quantity|total\s*qty|bill\s*(?:&|and)?\s*ship\s*(?:qty|quantity)|round\s*up|order\s*(?:qty|quantity)/i;
    }
        
    const chineseRegex = /[\u4e00-\u9fff]/;

    // Optimization: Jump to "sku information"
    let searchStartRow = 0;
    for (let i = 0; i < Math.min(data.length, 30); i++) {
        if (data[i] && data[i].join(" ").toLowerCase().includes("sku information")) {
            searchStartRow = i + 1;
            break;
        }
    }

    // 2. FIND HEADER
    const scanLimit = Math.min(data.length, 500);
    for (let i = searchStartRow; i < scanLimit; i++) {
        let row = data[i];
        if (!row || !Array.isArray(row)) continue;

        const rowText = row.join(" ").toLowerCase();
        const hasQty = row.some(cell => cell && qtyRegex.test(String(cell).trim()));

        if (customKeyword && hasQty) { startIndex = i; break; }

        if (rowText.includes("address:") || rowText.includes("attn:") || rowText.includes("country:")) continue;
        if (rowText.includes("tel#") || rowText.includes("email:") || rowText.includes("fax:")) continue;
        if (rowText.includes("just fill total qty") || rowText.includes("no moq")) continue;
        if (rowText.includes("round up") || rowText.includes("consider wastage")) continue;
        if (rowText.includes("refer to the chart") || rowText.includes("refer to chart")) continue;
        if (rowText.includes("kohls po quantities")|| rowText.includes("minimum")) continue; 
        if (rowText.includes("overrun") && rowText.includes("ordering qty")) continue;

        const hasChinese = row.some(cell => cell && chineseRegex.test(String(cell).trim()));

        if (hasQty && hasChinese) { startIndex = i; break; }
        if (rowText.includes("description10") || rowText.includes("description 10")) { startIndex = i; break; }
        if (hasQty) { startIndex = i; break; }
    }

    if (startIndex === -1) return [];

    // ==========================================
    // SMART MERGE LOGIC (Steps A, B, C, D)
    // ==========================================

    let headerRow = data[startIndex] || [];
    let nextRow = data[startIndex + 1]; 
    let dataStartIndex = startIndex + 1; 

    // Skip #N/A filler rows
    if (nextRow && Array.isArray(nextRow) && nextRow.length > 0 && nextRow.join(" ").toLowerCase().includes("#n/a")) {
        nextRow = data[startIndex + 2];
        dataStartIndex = startIndex + 2; 
    }

    let useCombinedHeader = false;
    if (nextRow && Array.isArray(nextRow) && nextRow.length > 0) {
        const subKeywords = ["UPC", "EAN", "FEATURE", "VSN", "PART", "OUT", "IN", "SIZE", "COLOR", "STYLE", "SKU"];
        useCombinedHeader = nextRow.some(c => {
            if (!c) return false;
            return subKeywords.some(k => String(c).toUpperCase().includes(k));
        });
        
        const hasBarcodeData = nextRow.some(c => /^\d{11,14}$/.test(String(c).trim()));
        if(hasBarcodeData) useCombinedHeader = false;

        // ⚡ NEW FIX: DATA ROW DETECTOR ⚡
        // If the next row has a number in the QTY column, it is DATA, not a header!
        const qtyIdx = headerRow.findIndex(c => c && qtyRegex.test(String(c).trim()));
        if (qtyIdx !== -1 && nextRow[qtyIdx] !== undefined && nextRow[qtyIdx] !== null) {
            const nextRowQtyVal = String(nextRow[qtyIdx]).trim().replace(/,/g, '');
            // Checks if the value is a number (e.g. 0, 2165)
            if (/^-?\d+(\.\d+)?$/.test(nextRowQtyVal)) {
                useCombinedHeader = false;
            }
        }
    }

    // --- STEP A: Build Unified Headers ---
    let unifiedHeaders = [];
    const maxLen = Math.max(headerRow.length || 0, (nextRow ? nextRow.length : 0));
    let lastMainHeader = "";

    for (let c = 0; c < maxLen; c++) {
        let val1 = headerRow[c] ? String(headerRow[c]).trim() : "";
        let val2 = (useCombinedHeader && nextRow && nextRow[c]) ? String(nextRow[c]).trim() : "";

        if (!val1 && lastMainHeader) val1 = lastMainHeader;
        else if (val1) lastMainHeader = val1;

        let finalName = val1;
        
        if (val2) {
            let v1Clean = val1.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
            let v2Clean = val2.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
            let v1Low = v1Clean.toLowerCase();
            let v2Low = v2Clean.toLowerCase();

            if (v1Low.includes("feature call") && v2Low.includes("feature icon")) {
                finalName = v2Clean;
            } else if (v2Low.includes("feature call") && v1Low.includes("feature icon")) {
                finalName = v1Clean;
            } else if (v2Clean.length <= 3 && !v2Low.includes("vsn")) {
                finalName = v1Clean + " " + v2Clean;
            } else if (v1Low.startsWith(v2Clean.split(/[\s_]+/)[0].toLowerCase())) {
                finalName = v2Clean;
            } else {
                finalName = v2Clean; 
            }
        }
        
        if(!finalName) finalName = "Column " + c;

        finalName = finalName.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
        unifiedHeaders.push({ name: finalName, originalIndex: c });
    }

    if (useCombinedHeader) dataStartIndex++;

    // --- STEP B: Extract Raw Data Body ---
    let rawBody = [];
    let emptyRowCount = 0;
    
    for (let i = dataStartIndex; i < data.length; i++) {
        let row = data[i];
        
        if (!row || row.every(c => !c || String(c).trim() === "")) {
            emptyRowCount++;
            // Bridges massive formatting gaps
            if (emptyRowCount >= 50) break;
            continue;
        }
        emptyRowCount = 0;

        const rowStr = row.join(" ").toLowerCase();

        if (rowStr.includes("total quantity") || 
            rowStr.includes("total qty") || 
            rowStr.includes("pls specify") || 
            rowStr.includes("working days") || 
            rowStr.includes("supplied by:") || 
            rowStr.includes("send order to email") || 
            rowStr.includes("information") || 
            rowStr.includes("factory as listed") || 
            rowStr.includes("south china contact") || 
            rowStr.includes("shipping instruction") || 
            (rowStr.includes("page") && rowStr.includes("of")) || 
            rowStr.includes("disclaimer") || 
            rowStr.startsWith("note") || 
            rowStr.startsWith("remarks") || 
            rowStr.includes("images")) {
            break;
        }

        const firstTextIndex = row.findIndex(c => c && String(c).trim().length > 0);
        if (firstTextIndex !== -1) {
            let firstText = String(row[firstTextIndex]).toLowerCase().trim();
            if (firstText === "total" || firstText.startsWith("total:") || firstText.startsWith("total qty") || firstText.startsWith("total quantity")) {
                break;
            }
        }

        if (rowStr.includes("please select") || rowStr.includes("#n/a") || rowStr.includes("#ref!") || (rowStr.includes("(max") && rowStr.includes("digits)"))) continue; 
        if (row[0] && String(row[0]).toLowerCase().trim().startsWith("eg")) continue;

        rawBody.push(row);
    }

    // --- STEP C: Smart Adjacent Merging ---
    let groupedHeaders = [];
    let currentGroup = null;

    for (let i = 0; i < unifiedHeaders.length; i++) {
        let col = unifiedHeaders[i];
        if (!currentGroup) {
            currentGroup = { name: col.name, indices: [col.originalIndex] };
        } else if (currentGroup.name === col.name) {
            currentGroup.indices.push(col.originalIndex); 
        } else {
            groupedHeaders.push(currentGroup);
            currentGroup = { name: col.name, indices: [col.originalIndex] }; 
        }
    }
    if (currentGroup) groupedHeaders.push(currentGroup);

    let finalIndices = [];
    let finalHeaderNames = [];

    // ⚡ FIX 2: RESTORE CLONE DETECTION (Replaces Winner Takes All) ⚡
    // This safely removes 100% identical duplicates
    // But keeps merged columns that contain DIFFERENT anomalies (like Womens_Size)
    groupedHeaders.forEach(group => {
        if (group.indices.length === 1) {
            finalIndices.push(group.indices[0]);
            finalHeaderNames.push(group.name);
        } else {
            let uniqueCols = []; 
            
            group.indices.forEach(idx => {
                let hasData = false;
                let columnData = [];
                
                rawBody.forEach(row => {
                    let val = (row[idx] === null || row[idx] === undefined) ? "" : String(row[idx]).trim();
                    columnData.push(val);
                    if (val !== "") hasData = true;
                });
                
                if (hasData) {
                    let isDuplicate = uniqueCols.some(keptIdx => {
                        let keptData = rawBody.map(row => (row[keptIdx] === null || row[keptIdx] === undefined) ? "" : String(row[keptIdx]).trim());
                        return columnData.every((val, r) => val === keptData[r]);
                    });
                    
                    if (!isDuplicate) {
                        uniqueCols.push(idx); 
                    }
                }
            });

            if (uniqueCols.length === 0) {
                finalIndices.push(group.indices[0]);
                finalHeaderNames.push(group.name);
            } else {
                uniqueCols.forEach(idx => {
                    finalIndices.push(idx);
                    finalHeaderNames.push(group.name);
                });
            }
        }
    });

    // Deduplicate Identical Columns (e.g. Size Field Gender 1, Size Field Gender 2)
    let duplicateTracker = {};
    finalHeaderNames.forEach(n => {
        duplicateTracker[n] = (duplicateTracker[n] || 0) + 1;
    });
    
    let currentCounts = {};
    for (let i = 0; i < finalHeaderNames.length; i++) {
        let name = finalHeaderNames[i];
        if (duplicateTracker[name] > 1) {
            currentCounts[name] = (currentCounts[name] || 0) + 1;
            finalHeaderNames[i] = name + " " + currentCounts[name];
        }
    }

    // --- STEP D: Build Final Table ---
    let cleanRows = [];
    cleanRows.push(finalHeaderNames); 

    rawBody.forEach(row => {
        let newRow = finalIndices.map(idx => {
            let val = row[idx];
            return (val === null || val === undefined) ? "" : val;
        });
        cleanRows.push(newRow);
    });

    let maxColsFinal = 0;
    cleanRows.forEach(r => { if (r.length > maxColsFinal) maxColsFinal = r.length; });
    cleanRows = cleanRows.map(r => {
        while (r.length < maxColsFinal) r.push("");
        return r;
    });

    return cleanRows;
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
        // SAFETY FIX: Ensure fileName exists before reading length
        const fName = p.fileName || ""; 

        if (fName !== lastFileName && fName !== "") {
            const shortName = (fName.length > 15) ? fName.substring(0, 12) + "..." : fName;
            list.innerHTML += `<div style="padding: 0 8px; display:flex; align-items:center; font-size:10px; color:#999; border-left:1px solid #ddd; margin-left:4px;">${shortName}</div>`;
            lastFileName = fName;
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
    
    // Initialize Counters
    let totalSets = projects.length;
    let processedSets = 0;
    let pendingSets = 0;
    
    let totalRows = 0; 
    let totalMatches = 0; 
    let totalMismatches = 0; 
    let tbody = "";

    projects.forEach((p, idx) => {
        // Count Status
        if (p.status === 'done') processedSets++;
        else pendingSets++;

        // Count Data Stats
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
                <td><button onclick="switchProject(${idx})" class="btn-ghost">View</button></td>
            </tr>`;
    });

    document.getElementById('overviewBody').innerHTML = tbody;
    
    // Updated Statistics Grid with MATCHES added
    document.getElementById('overviewStats').innerHTML = `
        <div class="big-stat">
            <div class="bs-val" style="color:#3b82f6">${totalSets}</div>
            <div class="bs-lbl">Total Sets</div>
        </div>
        <div class="big-stat">
            <div class="bs-val" style="color:#10b981">${processedSets}</div>
            <div class="bs-lbl">Processed</div>
        </div>
        <div class="big-stat">
            <div class="bs-val" style="color:#f59e0b">${pendingSets}</div>
            <div class="bs-lbl">To Process</div>
        </div>
        <div class="big-stat">
            <div class="bs-val" style="color:#10b981">${totalMatches}</div>
            <div class="bs-lbl">Total Matches</div>
        </div>
        <div class="big-stat">
            <div class="bs-val" style="color:#ef4444">${totalMismatches}</div>
            <div class="bs-lbl">Total Issues</div>
        </div>`;
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
    const undoBtn = document.getElementById('btnUndoRow');
    if (undoBtn) undoBtn.style.display = 'none';

    if (typeof excelState !== 'undefined') {
        excelState = { side: null, mode: null, r: -1, c: -1, editing: false };
    }
    if (typeof excelSelStart !== 'undefined') {
        excelSelStart = null; 
        excelSelEnd = null;
        isExcelDragging = false;
    }
    if (typeof currentSelection !== 'undefined') {
        currentSelection = { type: null, cells: [] };
    }
    
    loadProjectIntoView(idx);
    renderTopBar();
}

function saveCurrentViewToProject() {
    if (projects.length === 0 || isOverviewMode || activeProjectIdx < 0) return;
    
    const p = projects[activeProjectIdx];
    // Safety check: initialize settings if it's an old set
    if (!p.settings) p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };
    
    if (document.getElementById('step1').style.display !== 'none') {
        p.rawA = document.getElementById('tableA').value;
        p.rawB = document.getElementById('tableB').value;
        p.rawMatrix = document.getElementById('matrixRawInput').value;
        p.matrix = getMatrixDataFromUI(); 
        const matSec = document.getElementById('matrixSection');
        p.showMatrix = (matSec && matSec.style.display !== 'none');

        // Save Header Input Modes
        p.settings.modeA = document.getElementById('headerModeA')?.value || '1row';
        p.settings.modeB = document.getElementById('headerModeB')?.value || '1row';
    }

    if (document.getElementById('step4').style.display !== 'none') {
        // Save Checkbox States for Step 4
        p.settings.trimResults = document.getElementById('chkTrimResults')?.checked || false;
        p.settings.showErrors = document.getElementById('chkShowErrors')?.checked || false;
    }
}


function loadProjectIntoView(idx) {
    const p = projects[idx];
    if(!p) return; 

    document.getElementById('tableA').value = p.rawA || "";
    document.getElementById('tableB').value = p.rawB || "";
    
    let badge = document.getElementById('table2FileName');
    if (p.fileNameB) {
        if (!badge) {
            const table2Container = document.getElementById('tableB').parentElement;
            badge = document.createElement('div');
            badge.id = 'table2FileName';
            badge.style.cssText = "background:#e0f2fe; color:#0369a1; padding:5px 10px; font-size:12px; border-radius:4px; margin-bottom:5px; border:1px solid #bae6fd; display:inline-block;";
            table2Container.insertBefore(badge, document.getElementById('tableB'));
        }
        badge.innerHTML = `<strong>File:</strong> ${p.fileNameB}`;
        badge.style.display = 'inline-block';
    } else {
        if (badge) badge.style.display = 'none';
    }

    document.getElementById('previewTableA').innerHTML = "";
    document.getElementById('previewTableB').innerHTML = "";
    document.getElementById('countA').innerText = "0";
    document.getElementById('countB').innerText = "0";
    document.getElementById('mappingBody').innerHTML = "";
    
    const matrixInput = document.getElementById('matrixRawInput');
    if (matrixInput) {
        matrixInput.value = p.rawMatrix || "";
        matrixInput.oninput = function() { handleMatrixInput(this, idx); };
        matrixInput.onpaste = function(e) { handleMatrixPaste(e, idx); };
    }

    const matrixList = document.getElementById('matrixList');
    if (matrixList) {
        matrixList.innerHTML = "";
        if (p.matrix && p.matrix.length > 0) {
            p.matrix.forEach(m => addMatrixRow(m.key, m.val));
        } else if (p.rawMatrix) {
            p.matrix = parseMatrixString(p.rawMatrix); 
            if (p.matrix.length > 0) {
                 p.matrix.forEach(m => addMatrixRow(m.key, m.val));
            } else {
                 addMatrixRow();
            }
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

    // --- ENHANCEMENT: RESTORE INDEPENDENT SETTINGS TO THE UI ---
    if (!p.settings) p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };
    
    if (typeof setMode === 'function') {
        setMode('A', p.settings.modeA || '1row');
        setMode('B', p.settings.modeB || '1row');
    }

    const chkTrim = document.getElementById('chkTrimResults');
    const chkErr = document.getElementById('chkShowErrors');
    if (chkTrim) chkTrim.checked = p.settings.trimResults || false;
    if (chkErr) {
        chkErr.checked = p.settings.showErrors || false;
        toggleMismatchView(); // Apply the CSS hiding instantly
    }
    // -----------------------------------------------------------
    
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
    p.fileNameB = null; 
    
    // --- ENHANCEMENT: Reset Settings ---
    p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };

    const badge = document.getElementById('table2FileName');
    if (badge) badge.style.display = 'none';
    
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
// 8. DATA IMPORT (CPQ Excel)
// ==========================================

function handleCPQUpload(input) {
    if (!input.files || input.files.length === 0) return;

    saveCurrentViewToProject();

    // Sort files by Download/Modified time so they map predictably
    const files = Array.from(input.files).sort((a, b) => a.lastModified - b.lastModified);
    
    let processedCount = 0;
    const startIdx = activeProjectIdx;

    const processNext = (i) => {
        if (i >= files.length) {
            input.value = ""; 
            showModal("Upload Complete", `Successfully loaded <strong>${processedCount}</strong> CPQ files.`, "success");
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

        // Store filename in project
        projects[targetIdx].fileNameB = file.name;

        // Update Filename Badge dynamically if it's the currently active view
        if (targetIdx === activeProjectIdx) {
            const table2Container = document.getElementById('tableB').parentElement;
            let badge = document.getElementById('table2FileName');
            if (!badge) {
                badge = document.createElement('div');
                badge.id = 'table2FileName';
                badge.style.cssText = "background:#e0f2fe; color:#0369a1; padding:5px 10px; font-size:12px; border-radius:4px; margin-bottom:5px; border:1px solid #bae6fd; display:inline-block;";
                table2Container.insertBefore(badge, document.getElementById('tableB'));
            }
            badge.innerHTML = `<strong>File:</strong> ${file.name}`;
            badge.style.display = 'inline-block';
        }

        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                let extracted = "";
                
                // Smart Extract from Sheet 2 if exists
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
                
                // Fallback to Sheet 1
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
                processNext(i + 1); // Process next file
            }
        };
        reader.readAsArrayBuffer(file);
    };

    processNext(0); // Kickoff the loop
}

// ==========================================
// 9. PREVIEW & CLEANING
// ==========================================
function goToPreview() {
    try {
        // 1. Get current values
        const inputA = document.getElementById('tableA');
        const inputB = document.getElementById('tableB');
        
        // 2. Validation: Check if textareas are empty
        if (!inputA.value.trim() || !inputB.value.trim()) {
            showModal("Missing Data", "Please paste data into both tables.", 'error');
            return;
        }
        
        // 3. Update Project Memory
        const p = projects[activeProjectIdx];
        p.rawA = inputA.value; 
        p.rawB = inputB.value;
        
        const modeA = document.getElementById('headerModeA').value || "1row";
        const modeB = document.getElementById('headerModeB').value || "1row";
        
        // 4. Parse Data
        p.dataA = parseExcelData(p.rawA, modeA);
        p.dataB = parseExcelData(p.rawB, modeB);
        
        if (!p.dataA || !p.dataB) {
            showModal("Error Parsing", "Could not parse the data. Please check your input format.", 'error');
            return;
        }
        
        // 5. Clean Data (The new logic)
        autoCleanData(p.dataA);
        autoCleanData(p.dataB);
        
        // 6. Move to next step
        p.status = 'ready'; 
        p.step = 2;
        renderTopBar(); 
        renderPreviewTables(); 
        jumpToStep(2);
        
    } catch (err) {
        console.error("Preview Error:", err);
        showModal("System Error", "An error occurred while processing data:\n" + err.message, 'error');
    }
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

// --- GLOBAL EXCEL STATE ---
let excelState = {
    side: null,      // 'A' or 'B'
    mode: null,      // 'cell' or 'col'
    r: -1,           // Active Row Index
    c: -1,           // Active Column Index
    editing: false   // Is user currently typing?
};

function renderSinglePreview(containerId, countId, data, side) {
    if (!data || !data.body) return;

    document.getElementById(countId).innerText = data.body.length;
    const container = document.getElementById(containerId);
    
    // Add "Range Delete" Tools
    const tableBox = container.closest('.table-box');
    const header = tableBox.querySelector('.box-header');
    if (!header.querySelector('.custom-header-tools')) {
         let toolsDiv = document.createElement('div');
         toolsDiv.className = 'custom-header-tools';
         toolsDiv.style.display = 'flex'; toolsDiv.style.gap = '5px'; toolsDiv.style.alignItems = 'center';
         toolsDiv.innerHTML = `<input type="text" id="rangeInput${side}" placeholder="e.g. 2-5" style="padding:4px 8px; border-radius:4px; border:none; font-size:12px; color:#333; width:80px; outline:none;" onclick="event.stopPropagation()"> <button onclick="deleteRange('${side}'); event.stopPropagation()" style="background:rgba(255,255,255,0.2); color:white; border:1px solid rgba(255,255,255,0.4); padding:4px 10px; border-radius:4px; cursor:pointer; font-size:11px; font-weight:600;"><i class="fas fa-trash-alt"></i> Range</button>`;
         if (header.style.justifyContent !== 'space-between') { header.style.display = 'flex'; header.style.justifyContent = 'space-between'; header.style.alignItems = 'center'; }
         header.appendChild(toolsDiv);
    }

    const p = projects[activeProjectIdx];
    const isShowingHidden = p.uiState ? (side === 'A' ? p.uiState.showHiddenA : p.uiState.showHiddenB) : false;
    let displayList = [];
    
    data.body.forEach((row, idx) => displayList.push({ row, type: 'active', realIndex: idx, originalId: row._originalIdx ?? idx }));
    if (isShowingHidden && data.hiddenRows) {
        data.hiddenRows.forEach((item, idx) => displayList.push({ row: item.data, type: 'hidden', realIndex: idx, originalId: item.restoreIdx }));
    }
    displayList.sort((a, b) => a.originalId - b.originalId);

    const hiddenCount = data.hiddenRows ? data.hiddenRows.length : 0;
    const btnText = isShowingHidden ? "Hide Ignored" : `Show Ignored (${hiddenCount})`;
    const btnIcon = isShowingHidden ? "fa-eye-slash" : "fa-eye";
    const btnColor = isShowingHidden ? "#3b82f6" : "#64748b";
    const stickyStyle = "padding: 8px 10px; display:flex; justify-content:space-between; align-items:center; position: sticky; left: 0; background: #fff; z-index: 20; border-bottom: 1px solid #f1f5f9; width: fit-content; min-width:100%; border-right: 1px solid #e2e8f0; box-shadow: 2px 0 5px rgba(0,0,0,0.05);";

    let html = `
    <div style="${stickyStyle}">
        <button onclick="toggleHiddenTable('hiddenBox-${side}', this)" style="background:none; border:none; color:${btnColor}; font-size:12px; cursor:pointer; font-weight:600;"><i class="fas ${btnIcon}"></i> ${btnText}</button>
        <div style="display:flex; align-items:center; gap:10px;">
            <span style="font-size:11px; color:#64748b; font-weight:600; display:none;" id="undoBadge${side}"><i class="fas fa-undo"></i> Ctrl+Z to Undo</span>
            <button onclick="addNewRow('${side}')" style="background:#10b981; color:white; border:none; padding:4px 10px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;"><i class="fas fa-plus"></i> Add Row</button>
        </div>
    </div>
    
    <div style="outline:none; position:relative;" tabindex="0" id="gridContainer-${side}">
        <table class="clean-table" id="previewTable${side}" style="user-select:none; outline:none;" tabindex="0">
            <thead>
                <tr>
                    <th style="width:50px; text-align:center;">Action</th>
                    <th style="width:40px;">#</th>`;
    
    data.headers.forEach((h, i) => {
        html += `<th class="clickable-head" data-c="${i}" title="Click to select entire column">${h}</th>`;
    });
    
    html += "</tr></thead><tbody>";
    
    displayList.forEach((item, rIdx) => {
        const trStyle = item.type === 'hidden' ? "background-color: #f1f5f9; color: #94a3b8;" : "";
        const actionBtn = item.type === 'active' 
            ? `<button class="btn-ghost" style="color:#ef4444;" onclick="deleteRow('${side}', ${item.realIndex})"><i class="fas fa-trash"></i></button>`
            : `<button class="btn-ghost" style="color:#10b981;" onclick="restoreHiddenRow('${side}', ${item.realIndex})"><i class="fas fa-plus-circle"></i></button>`;

        html += `<tr style="${trStyle}" data-r="${item.realIndex}">
            <td style="text-align:center;">${actionBtn}</td>
            <td class="row-num" style="font-size:11px; color:#aaa; text-align:center; cursor:pointer;" title="Click to select row">${item.originalId + 1}</td>`;
            
        item.row.forEach((cell, cIdx) => {
             html += `<td data-r="${item.realIndex}" data-c="${cIdx}">${cell}</td>`;
        });
        html += `</tr>`;
    });

    html += `</tbody></table>
             <div id="fillHandle${side}" class="fill-handle" style="display:none;"></div>
             </div>`; 
    
    container.innerHTML = html;
    setTimeout(() => { GridEngine.init(side); }, 50);
}

function deleteRow(side, rowIndex) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    
    const rowToDelete = dataObj.body[rowIndex];
    if (!dataObj.hiddenRows) dataObj.hiddenRows = [];

    // Save with the existing Original ID
    // If for some reason _originalIdx is missing, fallback to current index (rowIndex)
    const originalId = (rowToDelete._originalIdx !== undefined) ? rowToDelete._originalIdx : rowIndex;

    dataObj.hiddenRows.unshift({ data: rowToDelete, restoreIdx: originalId });
    dataObj.body.splice(rowIndex, 1);
    
    renderPreviewTables();
}

function deleteRange(side) {
    const input = document.getElementById(`rangeInput${side}`);
    const val = input.value.trim();
    if (!val) return;
    
    const parts = val.split('-');
    if (parts.length !== 2) { 
        showModal("Invalid Format", "Use Start-End (e.g. 4-7)", "error"); return; 
    }
    
    let start = parseInt(parts[0]);
    let end = parseInt(parts[1]);
    if (isNaN(start) || isNaN(end) || start > end) return;
    
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    if (!dataObj.hiddenRows) dataObj.hiddenRows = [];
    
    let moved = 0;
    // Loop backwards
    for (let i = end; i >= start; i--) {
        const idx = i - 1; 
        if (idx >= 0 && idx < dataObj.body.length) {
            const row = dataObj.body[idx];
            const originalId = (row._originalIdx !== undefined) ? row._originalIdx : idx;
            
            dataObj.hiddenRows.unshift({ data: row, restoreIdx: originalId });
            dataObj.body.splice(idx, 1);
            moved++;
        }
    }
    
    if (moved > 0) { 
        showToast(`Moved ${moved} rows to Hidden list.`); 
        renderPreviewTables(); 
    }
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
            
        if (["quantity", "totalqty", "qty", "billshipquantity", "roundup", "billshipqty", "orderqty", "orderquantity"].includes(clean)) return "QTY_GROUP";
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
    
    if (document.getElementById('step3').style.display !== 'none') {
        saveMappingFromUI();
    }
    
    let processed = 0;
    let mismatchCount = 0;
    let missingDataLog = []; 

    projects.forEach(p => {
        // --- ENHANCEMENT: READ INDEPENDENT SETTINGS PER SET ---
        if (!p.settings) p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };
        const modeA = p.settings.modeA;
        const modeB = p.settings.modeB;
        const doTrim = p.settings.trimResults;

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
                
                if (!p.mapping || p.mapping.length === 0) {
                    autoMapProject(p);
                }
                
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
        
        // PO Logic Removed

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

function arrayToTSV(data) {
    if (!data) return "";
    return data.map(row => 
        row.map(cell => (cell == null ? "" : String(cell).replace(/[\r\n]+/g, " ").trim()))
           .join("\t")
    ).join("\n");
}


function findHeaderRowIndex(data) {
    const qtyRegex = /qty|quantity/i;
    for (let i = 0; i < Math.min(data.length, 60); i++) {
        if (data[i] && data[i].some(cell => cell && qtyRegex.test(String(cell).trim()))) {
            return i;
        }
    }
    return -1;
}


// --- LIVE MATRIX HANDLERS (With UI Update) ---

function handleMatrixInput(textarea, index = activeProjectIdx) {
    const p = projects[index];
    if (!p) return;

    // 1. Update Data
    p.rawMatrix = textarea.value;
    p.matrix = parseMatrixString(p.rawMatrix);

    // 2. UPDATE UI IMMEDIATELY (The missing link!)
    const listContainer = document.getElementById('matrixList');
    if (listContainer) {
        listContainer.innerHTML = ""; // Clear old rules
        p.matrix.forEach(m => addMatrixRow(m.key, m.val)); // Show new rules
    }
}

function handleMatrixPaste(e, index = activeProjectIdx) {
    e.preventDefault(); 
    let paste = (e.clipboardData || window.clipboardData).getData('text');
    
    // Clean messy text
    paste = paste
        .replace(/ need or not \(pls select\)/gi, "")
        .replace(/ \(Mandatory\)/gi, "")
        .replace(/ \(leave it blank if no need\)/gi, "")
        .replace(/ Field/gi, "")
        .replace(/  +/g, " ");

    const textarea = e.target;
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const originalText = textarea.value;
    
    const newText = originalText.substring(0, start) + paste + originalText.substring(end);
    
    // Update Text Box
    textarea.value = newText;
    
    // Trigger the UI Update
    handleMatrixInput(textarea, index);
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

    // 1. Initialize hidden bucket for rows that don't meet criteria
    data.hiddenRows = []; 

    // 2. Clean Headers: Remove empty or placeholder headers
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

    // 3. Tag every row with an Original ID for the Restore feature
    data.body.forEach((row, i) => {
        if (!row.hasOwnProperty('_originalIdx')) {
            Object.defineProperty(row, '_originalIdx', {
                value: i,
                writable: true,
                enumerable: false
            });
        }
    });

    // 4. Identify Quantity Columns using your specific Regex
    const qtyRegex = /qty|quantity|shipped|billed|units|pcs/i;
    const qtyIndices = [];
    data.headers.forEach((h, i) => { 
        if (h && qtyRegex.test(h.toString().toLowerCase())) qtyIndices.push(i); 
    });
    
    if (qtyIndices.length === 0) return;
    
    // 5. Categorize Rows into Active vs. Hidden
    const activeRows = [];

    data.body.forEach(row => {
        if (!row) return;
        const rowStr = row.join(" ").toLowerCase();

        // A. Your logic: Check if row is empty
        const isEmpty = !row.some(cell => cell && cell.toString().trim() !== "");
        
        // B. Your logic: Skip rows that look like repeated headers
        const isHeaderRepeat = rowStr.includes("upc") && rowStr.includes("style");

        // C. Your logic: Check for valid Quantity > 0
        let hasValidQty = false;
        for (let idx of qtyIndices) {
            let valStr = String(row[idx] || "").toLowerCase().replace(/[, \s]/g, '').replace(/pcs/g, '');
            let numVal = parseFloat(valStr);
            if (!isNaN(numVal) && numVal > 0) {
                hasValidQty = true;
                break;
            }
        }

        // Decision: If it fails your criteria, move it to the Hidden bucket
        if (isEmpty || isHeaderRepeat || !hasValidQty) {
            data.hiddenRows.push({ 
                data: row, 
                restoreIdx: row._originalIdx 
            });
        } else {
            activeRows.push(row);
        }
    });

    // Update the main body with only the "cleaned" active rows
    data.body = activeRows;
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


// --- HIDDEN ROW HANDLERS ---

function toggleHiddenTable(elementId, btn) {
    // Determine side (A or B)
    const side = elementId.includes('A') ? 'A' : 'B';
    const p = projects[activeProjectIdx];

    // Initialize UI state if missing
    if (!p.uiState) p.uiState = { showHiddenA: false, showHiddenB: false };

    // Toggle the state
    if (side === 'A') {
        p.uiState.showHiddenA = !p.uiState.showHiddenA;
    } else {
        p.uiState.showHiddenB = !p.uiState.showHiddenB;
    }

    // Refresh the table to show the new view
    renderPreviewTables();
}

function restoreHiddenRow(side, index) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    
    if (!dataObj || !dataObj.hiddenRows || !dataObj.hiddenRows[index]) return;

    // 1. Get row info
    const hiddenItem = dataObj.hiddenRows[index];
    const rowData = hiddenItem.data;
    const originalId = hiddenItem.restoreIdx;

    // 2. Find the correct insertion spot
    // We look for the first row in the active body that has a HIGHER original ID.
    // We insert our row *before* that one.
    let insertAt = dataObj.body.findIndex(r => (r._originalIdx !== undefined ? r._originalIdx : 999999) > originalId);

    // If we didn't find a higher ID, it means this row belongs at the very end.
    if (insertAt === -1) {
        insertAt = dataObj.body.length;
    }

    // 3. Insert Sequence
    dataObj.body.splice(insertAt, 0, rowData);

    // 4. Remove from hidden list
    dataObj.hiddenRows.splice(index, 1);

    // 5. Re-render
    renderPreviewTables();
}
// --- DATA EDITING HANDLERS ---

function updateCell(side, type, rowIndex, colIndex, cellElement) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    let newVal = cellElement.innerText.replace(/[\r\n]+/g, " ").trim();

    // 1. Update the data model first
    if (type === 'active') {
        if (dataObj && dataObj.body && dataObj.body[rowIndex]) {
            dataObj.body[rowIndex][colIndex] = newVal;
        }
    } else if (type === 'hidden') {
        if (dataObj && dataObj.hiddenRows && dataObj.hiddenRows[rowIndex]) {
            dataObj.hiddenRows[rowIndex].data[colIndex] = newVal;
        }
    }

    // 2. AUTOMATIC MOVE LOGIC
    // Check if this column is a "Quantity" column
    const headerName = dataObj.headers[colIndex] || "";
    const isQtyColumn = /^(qty|quantity|total\s*qty|total\s*quantity|units|pcs|bill|ship)$/i.test(headerName.trim());

    if (isQtyColumn) {
        // Parse the new value
        const cleanVal = newVal.replace(/,/g, '');
        const numVal = parseFloat(cleanVal);
        const isValid = !isNaN(numVal) && numVal > 0 && newVal !== "";

        if (type === 'active' && !isValid) {
            // Case: User cleared Quantity or set to 0 in an Active Row -> MOVE TO HIDDEN
            const row = dataObj.body[rowIndex];
            
            // Ensure hiddenRows array exists
            if (!dataObj.hiddenRows) dataObj.hiddenRows = [];
            
            // Add to hidden (Wrap it with restoreIdx logic)
            dataObj.hiddenRows.push({
                data: row,
                restoreIdx: row._originalIdx ?? 999999
            });
            
            // Remove from active
            dataObj.body.splice(rowIndex, 1);
            
            // Re-render immediately to reflect change
            renderPreviewTables();

        } else if (type === 'hidden' && isValid) {
            // Case: User entered a valid number in a Hidden Row -> MOVE TO ACTIVE
            const hiddenItem = dataObj.hiddenRows[rowIndex];
            const row = hiddenItem.data;
            
            // Add to active
            dataObj.body.push(row);
            
            // Remove from hidden
            dataObj.hiddenRows.splice(rowIndex, 1);
            
            // Re-render immediately
            renderPreviewTables();
        }
    }
}

function updateHiddenCell(side, rowIndex, colIndex, cellElement) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    
    // Access .data inside the hiddenRow object
    if (dataObj && dataObj.hiddenRows && dataObj.hiddenRows[rowIndex]) {
        let newVal = cellElement.innerText.replace(/[\r\n]+/g, " ").trim();
        dataObj.hiddenRows[rowIndex].data[colIndex] = newVal;
    }
}

// ==========================================
// NEW: MANUAL TABLE SELECTOR (CELL SELECTION)
// ==========================================

// Global Selection Variables for Manual Mapper
let selectionStartCol = -1;
let selectionEndCol = -1;

function reselectCurrentTable() {
    // 1. Safety Check
    if (!lastUploadedWorkbook) {
        alert("No file is currently loaded in memory. Please upload a file first.");
        return;
    }

    const p = projects[activeProjectIdx];
    const availableSheets = lastUploadedWorkbook.SheetNames;
    let targetSheet = availableSheets[0]; // Default to first sheet

    if (p) {
        // PRIORITY 1: Check if we stored the original sheet name
        if (p.originalSheetName && availableSheets.includes(p.originalSheetName)) {
            targetSheet = p.originalSheetName;
        } 
        // PRIORITY 2: Check if the Set Name matches a Sheet Name (Exact Match)
        // (e.g., if Set Name is "Sheet3", open "Sheet3")
        else if (availableSheets.includes(p.name)) {
            targetSheet = p.name;
        }
    }

    // 2. Open the mapper with the correct sheet pre-selected
    openManualMapper(lastUploadedWorkbook, lastUploadedFilename, targetSheet);
}


function openManualMapper(workbook, filename, targetSheetName = null) {
    manualWorkbook = workbook;
    manualFilename = filename;
    
    const select = document.getElementById('manualSheetSelect');
    select.innerHTML = "";
    if (workbook.SheetNames.length === 0) return alert("Empty Excel");

    let firstVisibleSheet = null;

    workbook.SheetNames.forEach(name => {
        // --- ENHANCEMENT: SKIP HIDDEN SHEETS ---
        if (workbook.Workbook && workbook.Workbook.Sheets) {
            const sMeta = workbook.Workbook.Sheets.find(s => s.name === name);
            if (sMeta && (sMeta.Hidden !== 0 || sMeta.state === 'hidden')) return;
        }

        if (!firstVisibleSheet) firstVisibleSheet = name;

        const opt = document.createElement('option');
        opt.value = name; 
        opt.innerText = name;
        if (targetSheetName && name === targetSheetName) {
            opt.selected = true;
        }
        select.appendChild(opt);
    });

    document.getElementById('manualSelectModal').style.display = 'block';
    
    // Load target sheet, or fallback to the first VISIBLE sheet
    const sheetToLoad = targetSheetName || firstVisibleSheet || workbook.SheetNames[0];
    currentSheetName = sheetToLoad; 
    renderRawSheet(sheetToLoad);
}

function renderRawSheet(sheetName) {
    currentSheetName = sheetName;
    
    // Reset selection state
    selectionStartRow = -1; selectionEndRow = -1;
    selectionStartCol = -1; selectionEndCol = -1;
    if (typeof updateManualButtonState === "function") updateManualButtonState();

    const sheet = manualWorkbook.Sheets[sheetName];
    
    // Use header:1 to get raw array of arrays
    const data = getVisibleExcelData(sheet);
    
    const table = document.getElementById('manualRawTable');
    table.innerHTML = "";

    // --- NEW: Make table focusable ---
    table.setAttribute('tabindex', '0');
    table.style.outline = 'none';

    // Render max 200 rows for performance
    const maxRows = Math.min(data.length, 200); 

    for(let r = 0; r < maxRows; r++) {
        const row = data[r];
        const tr = document.createElement('tr');
        
        // Row Header (Index Number)
        let html = `<td class="excel-idx" style="background:#f1f5f9; text-align:center; color:#888; font-size:10px; user-select:none;">${r+1}</td>`;
        
        for(let c = 0; c < row.length; c++) {
            // --- NEW: Clean ID and removed inline events (handled by helper now) ---
            html += `<td id="cell-${r}-${c}" 
                        data-r="${r}" data-c="${c}"
                        style="padding:6px; border:1px solid #e2e8f0; cursor:cell; min-width:50px; overflow:hidden; white-space:nowrap; max-width:200px; user-select: none;">
                        ${row[c] || ""}
                      </td>`;
        }
        tr.innerHTML = html;
        table.appendChild(tr);
    }

    // --- CRITICAL: ACTIVATE EXCEL FEATURES ---
    // This connects the table to the helper function you added at the bottom
    setTimeout(() => {
        enableExcelFeatures('manualRawTable'); 
    }, 100);
}

function handleCellMouseDown(el) {
    isDragging = true;
    selectionStartRow = parseInt(el.getAttribute('data-r'));
    selectionStartCol = parseInt(el.getAttribute('data-c'));
    
    selectionEndRow = selectionStartRow;
    selectionEndCol = selectionStartCol;
    
    highlightSelection();
    updateManualButtonState();

    // NEW: Start tracking mouse for auto-scroll
    document.addEventListener('mousemove', handleDragAutoScroll);
}

function handleCellMouseOver(el) {
    if (!isDragging) return;
    selectionEndRow = parseInt(el.getAttribute('data-r'));
    selectionEndCol = parseInt(el.getAttribute('data-c'));
    highlightSelection();
}

function handleCellMouseUp(el) {
    isDragging = false;
    updateManualButtonState();
}

function highlightSelection() {
    // 1. Clear previous highlights AND Reset Borders
    const cells = document.getElementById('manualRawTable').querySelectorAll('td');
    cells.forEach(el => {
        if(el.classList.contains('excel-idx')) return;
        el.style.backgroundColor = "transparent";
        el.style.color = "inherit";
        el.style.borderColor = "#e2e8f0"; // <--- Reset to default gray border
    });

    if (selectionStartRow === -1) return;

    // 2. Calculate bounds
    const rMin = Math.min(selectionStartRow, selectionEndRow);
    const rMax = Math.max(selectionStartRow, selectionEndRow);
    const cMin = Math.min(selectionStartCol, selectionEndCol);
    const cMax = Math.max(selectionStartCol, selectionEndCol);

    // 3. Highlight range
    for (let r = rMin; r <= rMax; r++) {
        for (let c = cMin; c <= cMax; c++) {
            const cell = document.getElementById(`cell-${r}-${c}`);
            if (cell) {
                // Make background AND border the same color for a solid look
                cell.style.backgroundColor = "#e0f2fe"; 
                cell.style.borderColor = "#e0f2fe"; // <--- Hides the white gap
                
                if (r === rMin) {
                    cell.style.backgroundColor = "#2563eb"; 
                    cell.style.borderColor = "#2563eb"; // <--- Hides gap in header
                    cell.style.color = "white";
                }
            }
        }
    }
}

function updateManualButtonState() {
    const btn = document.getElementById('btnManualConfirm');
    const msg = document.getElementById('manualInstruction');
    
    if (selectionStartRow !== -1 && selectionEndRow !== -1) {
        const rows = Math.abs(selectionEndRow - selectionStartRow) + 1;
        const cols = Math.abs(selectionEndCol - selectionStartCol) + 1;
        
        btn.disabled = false;
        btn.style.background = "#2563eb";
        btn.style.cursor = "pointer";
        msg.innerHTML = `<i class="fas fa-check-circle"></i> Selected: <strong>${rows} Rows</strong> x <strong>${cols} Columns</strong>.`;
        msg.style.color = "#166534";
    } else {
        btn.disabled = true;
        btn.style.background = "#cbd5e1";
        btn.style.cursor = "not-allowed";
        msg.innerHTML = `<i class="fas fa-mouse-pointer"></i> <strong>CLICK & DRAG</strong> to select header + data range.`;
        msg.style.color = "#ef4444";
    }
}


function confirmManualImport() {
    if (selectionStartRow === -1) return;

    try {
        // 1. Get Data from Workbook
        const sheet = manualWorkbook.Sheets[currentSheetName];
        const rawData = getVisibleExcelData(sheet);
        
        // 2. Slice Data based on selection
        const rMin = Math.min(selectionStartRow, selectionEndRow);
        const rMax = Math.max(selectionStartRow, selectionEndRow);
        const cMin = Math.min(selectionStartCol, selectionEndCol);
        const cMax = Math.max(selectionStartCol, selectionEndCol);
        
        let slicedData = [];
        for(let r = rMin; r <= rMax; r++) {
            let newRow = [];
            const srcRow = rawData[r] || [];
            for(let c = cMin; c <= cMax; c++) {
                newRow.push(srcRow[c] || "");
            }
            slicedData.push(newRow);
        }

        if (slicedData.length < 2) {
            alert("Please select at least 2 rows (1 Header + 1 Data).");
            return;
        }

        const newRawData = arrayToTSV(slicedData);
        
        // ===============================================
        // NEW LOGIC: DETECT DUPLICATE & ASK USER
        // ===============================================
        
        // Check if a tab with this sheet name ALREADY exists
        const existingIdx = projects.findIndex(p => p.name === currentSheetName);

        if (existingIdx !== -1) {
            // CONFLICT FOUND: Ask the user what to do
            showConflictModal(currentSheetName, existingIdx, newRawData);
        } else {
            // NO CONFLICT: Create new immediately
            applyManualNew(newRawData);
        }

    } catch (err) {
        console.error(err);
        alert("Error importing: " + err.message);
    }
}

function closeManualModal() {
    document.getElementById('manualSelectModal').style.display = 'none';
    manualWorkbook = null;
    manualFilename = "";
}

// ==========================================
// MANUAL SELECTION AUTO-SCROLL LOGIC
// ==========================================
let scrollInterval = null;
let scrollVector = { x: 0, y: 0 };

function handleDragAutoScroll(e) {
    if (!isDragging) return;

    const table = document.getElementById('manualRawTable');
    if (!table) return;
    const container = table.parentElement; // The scrollable div wrapper
    if (!container) return;

    const rect = container.getBoundingClientRect();
    const buffer = 50; // How close to edge before scrolling starts (pixels)
    const speed = 20;  // How fast it scrolls

    let vx = 0;
    let vy = 0;

    // Check Right Edge
    if (e.clientX > rect.right - buffer) vx = speed;
    // Check Left Edge
    else if (e.clientX < rect.left + buffer) vx = -speed;
    
    // Check Bottom Edge
    if (e.clientY > rect.bottom - buffer) vy = speed;
    // Check Top Edge
    else if (e.clientY < rect.top + buffer) vy = -speed;

    scrollVector = { x: vx, y: vy };

    // Start or Stop Interval based on mouse position
    if (vx !== 0 || vy !== 0) {
        if (!scrollInterval) {
            scrollInterval = setInterval(() => {
                container.scrollLeft += scrollVector.x;
                container.scrollTop += scrollVector.y;
            }, 30); // Runs every 30ms
        }
    } else {
        stopAutoScroll();
    }
}

function stopAutoScroll() {
    clearInterval(scrollInterval);
    scrollInterval = null;
    scrollVector = { x: 0, y: 0 };
}

// ==========================================
// NEW: PASTE & ADD ROW LOGIC
// ==========================================

function addNewRow(side) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    
    // Create an empty row with the same number of columns as headers
    const colCount = dataObj.headers.length;
    const newRow = new Array(colCount).fill("");
    
    // Assign a new original ID (max existing + 1)
    let maxId = 0;
    if (dataObj.body.length > 0) maxId = Math.max(...dataObj.body.map(r => r._originalIdx || 0));
    if (dataObj.hiddenRows && dataObj.hiddenRows.length > 0) {
        const hiddenMax = Math.max(...dataObj.hiddenRows.map(h => h.restoreIdx || 0));
        maxId = Math.max(maxId, hiddenMax);
    }
    
    Object.defineProperty(newRow, '_originalIdx', { value: maxId + 1, writable: true, enumerable: false });
    
    // Add to body
    dataObj.body.push(newRow);
    
    renderPreviewTables();
    showToast("Row Added");
}

function handlePaste(e, side, rowIndex, colIndex) {
    // 1. Stop default paste (which puts all text in one cell)
    e.preventDefault();
    e.stopPropagation();

    // 2. Get Data from Clipboard
    const clipboardData = (e.clipboardData || window.clipboardData).getData('text');
    if (!clipboardData) return;

    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    const rows = dataObj.body;

    // 3. Split into Rows and Columns (Excel uses \n for rows, \t for cols)
    const pasteRows = clipboardData.split(/\r\n|\n|\r/);
    
    let rowsUpdated = 0;

    pasteRows.forEach((pasteRow, i) => {
        if (!pasteRow && i === pasteRows.length - 1) return; // Skip trailing newline
        
        const targetRowIdx = rowIndex + i;
        
        // Stop if we run out of rows in the table
        if (targetRowIdx >= rows.length) return; 

        const pasteCells = pasteRow.split('\t');
        
        pasteCells.forEach((cellValue, j) => {
            const targetColIdx = colIndex + j;
            
            // Stop if we run out of columns
            if (targetColIdx < rows[targetRowIdx].length) {
                // Update the data model directly
                rows[targetRowIdx][targetColIdx] = cellValue.trim();
            }
        });
        rowsUpdated++;
    });

    // 4. Refresh Table to show new data
    renderPreviewTables();
    showToast(`Pasted ${rowsUpdated} rows`);
}


// 4. Add New Row Function
function addNewRow(side) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    
    // Create empty row matching header length
    const colCount = dataObj.headers.length;
    const newRow = new Array(colCount).fill("");
    
    // Generate new ID
    let maxId = 0;
    if (dataObj.body.length > 0) maxId = Math.max(...dataObj.body.map(r => r._originalIdx || 0));
    if (dataObj.hiddenRows && dataObj.hiddenRows.length > 0) {
        const hiddenMax = Math.max(...dataObj.hiddenRows.map(h => h.restoreIdx || 0));
        maxId = Math.max(maxId, hiddenMax);
    }
    
    // Tag with ID
    Object.defineProperty(newRow, '_originalIdx', { value: maxId + 1, writable: true, enumerable: false });
    
    dataObj.body.push(newRow);
    renderPreviewTables();
    showToast("Row Added");
}

// 5. Excel-Style Copy/Paste Handler
function handlePaste(e, side, rowIndex, colIndex) {
    // If user is editing inside the cell (cursor blinking), let normal paste happen
    if (e.target.isContentEditable && e.target.getAttribute('contenteditable') === 'true') return;

    e.preventDefault();
    const clipboardData = (e.clipboardData || window.clipboardData).getData('text');
    if (!clipboardData) return;

    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    const rows = dataObj.body;

    const pasteRows = clipboardData.split(/\r\n|\n|\r/);
    let rowsUpdated = 0;

    pasteRows.forEach((pasteRow, i) => {
        if (!pasteRow && i === pasteRows.length - 1) return; 
        
        const targetRowIdx = rowIndex + i;
        if (targetRowIdx >= rows.length) return; 

        const pasteCells = pasteRow.split('\t');
        pasteCells.forEach((cellValue, j) => {
            const targetColIdx = colIndex + j;
            if (targetColIdx < rows[targetRowIdx].length) {
                rows[targetRowIdx][targetColIdx] = cellValue.trim();
            }
        });
        rowsUpdated++;
    });

    renderPreviewTables();
    showToast(`Pasted ${rowsUpdated} rows`);
}

// ==========================================
// HELPER FUNCTIONS (Paste at the bottom of script.js)
// ==========================================

function findHeaderRowIndex(data) {
    const qtyRegex = /qty|quantity/i;
    for (let i = 0; i < Math.min(data.length, 60); i++) {
        let row = data[i];
        if (!row) continue;
        if (row.some(cell => cell && qtyRegex.test(String(cell).trim()))) {
            return i;
        }
    }
    return -1;
}


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
        let cleanLine = line.trim(); 

        let k = "", v = "";
        
        // 1. Split by colon FIRST before deleting any parentheses
        if (cleanLine.includes(":")) {
            let idx = cleanLine.indexOf(":");
            k = cleanLine.substring(0, idx);
            v = cleanLine.substring(idx+1);
        } else {
            k = cleanLine; 
        }

        // 2. Clean the KEY: Remove quotes, parentheses, and extra spaces
        k = k.replace(/"/g, "").replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim();
        
        // 3. Clean the VALUE: Just trim whitespace. DO NOT remove parentheses!
        v = v ? v.trim() : "";

        // Standardize Known Keys
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

// ===============================================
// MANUAL IMPORT HELPERS
// ===============================================

function showConflictModal(name, existingIdx, newData) {
    const modal = document.getElementById('customModal');
    const titleEl = document.getElementById('modalTitle');
    const msgEl = document.getElementById('modalMsg');
    const iconEl = document.getElementById('modalIcon');
    
    const confirmBtn = document.getElementById('modalBtn');
    const cancelBtn = document.getElementById('modalCancelBtn'); 

    // 1. Setup UI Text
    titleEl.innerText = "Sheet Already Exists";
    msgEl.innerHTML = `The sheet <b>"${name}"</b> is already loaded.<br>Do you want to REPLACE the existing tab or Create a NEW one?`;
    
    iconEl.className = 'fas fa-question-circle sa-icon-warn';
    iconEl.style.color = '#f59e0b'; 

    // 2. Configure "REPLACE" (Primary Button)
    confirmBtn.innerText = "Replace Existing";
    confirmBtn.style.backgroundColor = '#f59e0b'; // Orange
    // Remove old listeners to prevent stacking
    const newConfirm = confirmBtn.cloneNode(true);
    confirmBtn.parentNode.replaceChild(newConfirm, confirmBtn);
    
    newConfirm.onclick = function() {
        applyManualReplace(existingIdx, newData);
        closeModal();
    };

    // 3. Configure "CREATE NEW" (Secondary Button)
    cancelBtn.style.display = 'inline-block';
    cancelBtn.innerText = "Create New Tab";
    cancelBtn.style.backgroundColor = '#2563eb'; // Blue
    cancelBtn.style.color = 'white';
    
    const newCancel = cancelBtn.cloneNode(true);
    cancelBtn.parentNode.replaceChild(newCancel, cancelBtn);

    newCancel.onclick = function() {
        applyManualNew(newData);
        closeModal();
    };

    modal.classList.add('open');
}

function applyManualReplace(idx, rawData) {
    const p = projects[idx];
    p.rawA = rawData;
    p.dataA = null; 
    p.status = 'ready';
    p.step = 1;     
    
    // 'true' skips the save step so we don't overwrite our new work
    switchProject(idx, true); 
    
    document.getElementById('manualSelectModal').style.display = 'none';
    showToast(`Replaced data in "${p.name}"`);
}

function applyManualNew(rawData) {
    let finalName = currentSheetName;
    let counter = 1;
    
    // Find unique name
    while(projects.some(p => p.name === finalName)) {
        counter++;
        finalName = `${currentSheetName} (${counter})`;
    }

    createSet(finalName, null, currentSheetName, manualFilename);
    const newIdx = projects.length - 1;
    
    projects[newIdx].rawA = rawData;
    projects[newIdx].status = 'ready';
    
    document.getElementById('manualSelectModal').style.display = 'none';
    renderTopBar();
    switchProject(newIdx, true); // Skip save here too just to be safe
    showToast(`Created "${finalName}"`);
}


// --- ADVANCED EXCEL LOGIC ENGINE (With Auto-Scroll) ---
let excelSelStart = null; 
let excelSelEnd = null;
let isExcelDragging = false;
let autoScrollTimer = null;
let lastMouseX = 0;
let lastMouseY = 0;

function enableExcelFeatures(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return;
    
    // 1. Identify the scrollable container (Parent of the table)
    const container = table.parentElement; 
    container.style.position = 'relative'; // Ensure relative positioning

    // --- MOUSE DOWN (Start Selection) ---
    table.onmousedown = function(e) {
        // Ignore right-clicks or inputs
        if (e.button !== 0 || ['INPUT', 'BUTTON', 'SELECT'].includes(e.target.tagName)) return;

        const cell = e.target.closest('td');
        if (!cell || cell.classList.contains('excel-idx')) return;

        isExcelDragging = true;
        const startR = parseInt(cell.getAttribute('data-r'));
        const c = parseInt(cell.getAttribute('data-c'));

        const maxR = table.rows.length - 1;
        const maxC = (table.rows[0] ? table.rows[0].cells.length - 2 : 20); 

        if (e.shiftKey && excelSelStart) {
            // Normal shift-click expansion
            excelSelEnd = { r: startR, c: c };
        } else {
            // --- ENHANCEMENT: THE SMART DOWNWARD SCANNER ---
            let endR = maxR;
            let emptyCount = 0;

            // Scan downwards from the row you just clicked
            for (let i = startR + 1; i <= maxR; i++) {
                const rowTr = table.rows[i];
                if (!rowTr) continue;

                // Extract row text
                let rowTextArray = [];
                let hasData = false;
                for (let col = 1; col < rowTr.cells.length; col++) { 
                    if(rowTr.cells[col].classList.contains('excel-idx')) continue;
                    
                    const cellText = rowTr.cells[col].innerText.trim().toLowerCase();
                    rowTextArray.push(cellText);
                    if (cellText) hasData = true;
                }
                const rowTextStr = rowTextArray.join(" ");

                // 1. Stop if we hit 5 completely blank rows in a row
                if (!hasData) {
                    emptyCount++;
                    if (emptyCount >= 5) { endR = i - 5; break; }
                    continue;
                } else {
                    emptyCount = 0;
                }

                // 2. Stop if we hit any of our known STOP WORDS
                if (
                    rowTextStr.includes("address:") || rowTextStr.includes("attn:") || rowTextStr.includes("country:") ||
                    rowTextStr.includes("tel#") || rowTextStr.includes("email:") || rowTextStr.includes("fax:") ||
                    rowTextStr.includes("just fill total qty") || rowTextStr.includes("no moq") ||
                    rowTextStr.includes("round up") || rowTextStr.includes("consider wastage") ||
                    rowTextStr.includes("refer to the chart") || rowTextStr.includes("refer to chart") ||
                    rowTextStr.includes("kohls po quantities") || rowTextStr.includes("minimum") ||
                    (rowTextStr.includes("overrun") && rowTextStr.includes("ordering qty")) ||
                    rowTextStr.includes("information") || rowTextStr.includes("factory as listed") ||
                    rowTextStr.includes("south china contact") || rowTextStr.includes("shipping instruction") ||
                    (rowTextStr.includes("page") && rowTextStr.includes("of")) ||
                    rowTextStr.includes("disclaimer") || rowTextStr.startsWith("note") || rowTextStr.startsWith("remarks") ||
                    rowTextStr.includes("images")
                ) {
                    endR = i - 1; // Stop at the row right before this one
                    break;
                }

                // 3. Stop if we hit the "TOTAL" row
                const firstText = rowTextArray.find(t => t.length > 0);
                if (firstText && (firstText === "total" || firstText.startsWith("total:") || firstText.startsWith("total qty"))) {
                    endR = i - 1;
                    break;
                }
            }

            // Apply the smart selection!
            if (endR < startR) endR = startR; // Safety check
            
            // FIX: Start the selection EXACTLY at the column you clicked, not column 0!
            excelSelStart = { r: startR, c: c }; 
            excelSelEnd = { r: endR, c: maxC };
        }

        // Update Global Variables (Legacy Support)
        updateLegacyGlobals(excelSelStart.r, excelSelStart.c, excelSelEnd.r, excelSelEnd.c);

        highlightExcelRange(table);
        table.focus(); // Capture keyboard focus

        // START Auto-Scroll Monitoring
        if (typeof startAutoScroll === 'function') startAutoScroll(container);
    };

    // --- MOUSE MOVE (Track Mouse & Update Selection) ---
    // We attach this to DOCUMENT so it works even if you drag OUTSIDE the table
    document.addEventListener('mousemove', function(e) {
        if (!isExcelDragging) return;

        // Update global mouse coordinates for the auto-scroller
        if (typeof lastMouseX !== 'undefined') lastMouseX = e.clientX;
        if (typeof lastMouseY !== 'undefined') lastMouseY = e.clientY;

        // 1. Find the cell under the cursor
        // We hide the cursor temporarily to see what element is strictly underneath
        const el = document.elementFromPoint(e.clientX, e.clientY);
        const cell = el ? el.closest('td') : null;

        if (cell && table.contains(cell)) {
            const r = parseInt(cell.getAttribute('data-r'));
            const c = parseInt(cell.getAttribute('data-c'));
            
            if (!isNaN(r) && !isNaN(c)) {
                excelSelEnd = { r, c };
                updateLegacyGlobals(null, null, r, c);
                highlightExcelRange(table);
            }
        }
    });

    // --- MOUSE UP (Stop Everything) ---
    document.addEventListener('mouseup', function() { 
        isExcelDragging = false; 
        if (typeof stopAutoScroll === 'function') stopAutoScroll();
    });

    // --- KEYBOARD EVENTS (Copy/Paste/Arrows) ---
    table.onkeydown = async function(e) {
        if (!excelSelStart) return;

        // COPY (Ctrl+C)
        if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
            e.preventDefault();
            copyExcelData(table);
        }

        // PASTE (Ctrl+V)
        if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
            e.preventDefault();
            try {
                const text = await navigator.clipboard.readText();
                pasteExcelData(table, text);
            } catch (err) {}
        }

        // ARROWS
        if (e.key.startsWith('Arrow')) {
            e.preventDefault();
            moveExcelSelection(e.key, table);
        }
    };
}

// --- HELPER: Auto-Scroll Logic ---
function startAutoScroll(container) {
    stopAutoScroll(); // Clear existing
    
    autoScrollTimer = setInterval(() => {
        if (!isExcelDragging) return;

        const rect = container.getBoundingClientRect();
        const sensitivity = 50; // Distance from edge (px) to trigger scroll
        const speed = 15;       // Scroll speed (px per tick)

        let scrolled = false;

        // Vertical Scroll
        if (lastMouseY < rect.top + sensitivity) {
            container.scrollTop -= speed;
            scrolled = true;
        } else if (lastMouseY > rect.bottom - sensitivity) {
            container.scrollTop += speed;
            scrolled = true;
        }

        // Horizontal Scroll
        if (lastMouseX < rect.left + sensitivity) {
            container.scrollLeft -= speed;
            scrolled = true;
        } else if (lastMouseX > rect.right - sensitivity) {
            container.scrollLeft += speed;
            scrolled = true;
        }

        // If we scrolled, we need to update the selection to the new cell under the cursor!
        if (scrolled) {
            // Re-trigger the selection update logic manually
            const el = document.elementFromPoint(lastMouseX, lastMouseY);
            const cell = el ? el.closest('td') : null;
            if (cell) {
                const r = parseInt(cell.getAttribute('data-r'));
                const c = parseInt(cell.getAttribute('data-c'));
                if (!isNaN(r) && !isNaN(c)) {
                    excelSelEnd = { r, c };
                    // Update global legacy vars
                    if (typeof selectionEndRow !== 'undefined') {
                        selectionEndRow = r; selectionEndCol = c;
                    }
                    // We need to find the table again to highlight
                    const table = cell.closest('table');
                    if(table) highlightExcelRange(table);
                }
            }
        }
    }, 30); // Run every 30ms
}

function stopAutoScroll() {
    if (autoScrollTimer) clearInterval(autoScrollTimer);
    autoScrollTimer = null;
}

// --- HELPER: Update Global Variables (Legacy Support) ---
function updateLegacyGlobals(startR, startC, endR, endC) {
    if (startR !== null && typeof selectionStartRow !== 'undefined') {
        selectionStartRow = startR;
        selectionStartCol = startC;
    }
    if (endR !== null && typeof selectionEndRow !== 'undefined') {
        selectionEndRow = endR;
        selectionEndCol = endC;
    }
    if (typeof updateManualButtonState === "function") updateManualButtonState();
}

// --- (Keep the highlightExcelRange, copyExcelData, pasteExcelData functions exactly as they were) ---
// ... (Make sure you include them here if you overwrote the whole block) ...

// --- VISUAL HIGHLIGHTER ---
function highlightExcelRange(table) {
    // Remove old highlights
    table.querySelectorAll('.excel-selected').forEach(el => el.classList.remove('excel-selected'));

    if (!excelSelStart || !excelSelEnd) return;

    // Math to find the square box
    const r1 = Math.min(excelSelStart.r, excelSelEnd.r);
    const r2 = Math.max(excelSelStart.r, excelSelEnd.r);
    const c1 = Math.min(excelSelStart.c, excelSelEnd.c);
    const c2 = Math.max(excelSelStart.c, excelSelEnd.c);

    // Apply class to all cells in range
    for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell) cell.classList.add('excel-selected');
        }
    }
}

// --- LOGIC: COPY ---
function copyExcelData(table) {
    const selected = table.querySelectorAll('.excel-selected');
    if (selected.length === 0) return;

    // Group cells by Row Index to build the TSV string
    let rowsMap = {};
    selected.forEach(cell => {
        const r = parseInt(cell.getAttribute('data-r'));
        if (!rowsMap[r]) rowsMap[r] = [];
        rowsMap[r].push(cell);
    });

    // Sort rows and columns to ensure correct order
    const sortedRows = Object.keys(rowsMap).sort((a,b) => a-b);
    const tsv = sortedRows.map(r => {
        const cells = rowsMap[r].sort((a,b) => a.getAttribute('data-c') - b.getAttribute('data-c'));
        return cells.map(c => c.innerText).join('\t');
    }).join('\t\n'); // Use tab for columns, newline for rows

    navigator.clipboard.writeText(tsv).then(() => {
        console.log("Copied to clipboard"); 
    });
}

// --- LOGIC: PASTE ---
function pasteExcelData(table, text) {
    if (!excelSelStart) return;
    
    const rows = text.split(/\r\n|\n|\r/).filter(r => r);
    const startR = Math.min(excelSelStart.r, excelSelEnd.r);
    const startC = Math.min(excelSelStart.c, excelSelEnd.c);

    rows.forEach((rowStr, rOffset) => {
        const cols = rowStr.split('\t');
        cols.forEach((val, cOffset) => {
            const targetR = startR + rOffset;
            const targetC = startC + cOffset;
            const cell = table.querySelector(`td[data-r="${targetR}"][data-c="${targetC}"]`);
            if (cell) {
                cell.innerText = val.trim();
                // Flash effect to show paste happened
                cell.style.backgroundColor = "#dcfce7"; 
                setTimeout(() => cell.style.backgroundColor = "", 300);
            }
        });
    });
}

// --- LOGIC: ARROW MOVEMENT ---
function moveExcelSelection(key, table) {
    let r = excelSelEnd.r; // Move from the ACTIVE end
    let c = excelSelEnd.c;

    if (key === 'ArrowUp') r--;
    if (key === 'ArrowDown') r++;
    if (key === 'ArrowLeft') c--;
    if (key === 'ArrowRight') c++;

    // Boundaries check
    const maxR = table.rows.length - 1; 
    const maxC = (table.rows[r]?.cells.length || 20) - 1; 

    if (r < 0) r = 0;
    if (c < 0) c = 0;
    
    // Update State
    excelSelStart = { r, c }; // Reset selection to single cell on move
    excelSelEnd = { r, c };
    
    // Update Globals
    updateLegacyGlobals(r, c, r, c);

    highlightExcelRange(table);
    
    // Auto-scroll to cell
    const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
    if (cell) cell.scrollIntoView({block: 'nearest', inline: 'nearest'});
}

//STEP-2 EDIITING HELPER FUNCTIONS
function handleHeaderEdit(side, colIdx, newName) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    if(data && data.headers) {
        data.headers[colIdx] = newName.trim();
    }
}

function handleCellEdit(side, rIdx, cIdx, newVal) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    if(data && data.body && data.body[rIdx]) {
        data.body[rIdx][cIdx] = newVal.trim();
    }
}

function handleColumnSelect(side, colIdx) {
    if (selectedColSide === side && selectedColIdx === colIdx) {
        resetColSelection();
    } else {
        selectedColSide = side;
        selectedColIdx = colIdx;
    }
    renderPreviewTables();
}

function resetColSelection() {
    selectedColSide = null;
    selectedColIdx = -1;
    const all = document.querySelectorAll('th, td');
    all.forEach(el => {
        el.style.backgroundColor = '';
        el.style.borderBottom = '';
    });
}

function handleSmartPaste(e, side) {
    if (selectedColSide !== side || selectedColIdx === -1) return;
    
    e.preventDefault();
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;

    const pasteRows = text.split(/\r\n|\n|\r/).filter(r => r.trim());
    if (pasteRows.length === 0) return;

    for (let i = 0; i < data.body.length; i++) {
        const pasteVal = pasteRows[i % pasteRows.length].split('\t')[0];
        data.body[i][selectedColIdx] = pasteVal;
    }
    
    renderPreviewTables();
}

function handleRowDragStart(e, side, rIdx) {
    dragSrcRow = rIdx;
    dragSrcSide = side;
    e.dataTransfer.effectAllowed = 'move';
    e.target.style.opacity = '0.5';
}

function handleRowDragOver(e) {
    if (e.preventDefault) e.preventDefault(); 
    e.dataTransfer.dropEffect = 'move';
    return false;
}

function handleRowDrop(e, side, targetRIdx) {
    e.stopPropagation();
    if (dragSrcSide !== side) return;
    
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    
    const item = data.body[dragSrcRow];
    data.body.splice(dragSrcRow, 1);
    data.body.splice(targetRIdx, 0, item);
    
    renderPreviewTables();
    return false;
}

// --- STEP 2: EXCEL FEATURES (Original Version) ---

// 1. SELECT SINGLE CELL
function selectExcelCell(side, r, c) {
    if (excelState.editing) return; 
    excelState = { side: side, mode: 'cell', r: r, c: c, editing: false };
    renderPreviewTables();
}

// 2. SELECT WHOLE COLUMN
function selectExcelCol(side, c) {
    if (excelState.side === side && excelState.mode === 'col' && excelState.c === c) {
        excelState = { side: null, mode: null, r: -1, c: -1, editing: false };
    } else {
        excelState = { side: side, mode: 'col', r: -1, c: c, editing: false };
    }
    renderPreviewTables();
}

// 3. EDIT CELL (Triggered by DblClick or Typing)
function editExcelCell(side, r, c, cellEl, initialValue = null) {
    excelState.editing = true;
    cellEl.contentEditable = true;
    cellEl.classList.add('excel-editing');
    
    // If triggered by typing a letter, overwrite immediately
    if (initialValue !== null) {
        cellEl.innerText = initialValue;
        // Move cursor to end
        const range = document.createRange();
        const sel = window.getSelection();
        range.selectNodeContents(cellEl);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
    }
    
    cellEl.focus();
    
    // Save on Blur
    cellEl.onblur = function() {
        saveExcelEdit(side, r, c, this.innerText);
    };
    
    // Save on Enter
    cellEl.onkeydown = function(e) {
        if(e.key === 'Enter') {
            e.preventDefault();
            this.blur(); // Triggers save
        }
    };
}

// 4. SAVE DATA
function saveExcelEdit(side, r, c, val) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    if (data && data.body[r]) {
        data.body[r][c] = val.trim();
    }
    excelState.editing = false;
    renderPreviewTables(); // Re-render to lock the cell back to static
}

// 5. KEYBOARD NAVIGATION & "TYPE TO EDIT"
function handleExcelKey(e, side) {
    // If we are currently typing inside a cell, let default behavior happen 
    // (except Enter/Esc/Tab which we want to control)
    if (excelState.editing) {
        if (e.key === 'Enter') { 
            e.preventDefault(); 
            document.activeElement.blur(); // Save & Exit
        }
        else if (e.key === 'Escape') {
            e.preventDefault();
            // Cancel edit: Re-render table to revert value
            excelState.editing = false;
            renderPreviewTables(); 
        }
        else if (e.key === 'Tab') {
            e.preventDefault();
            document.activeElement.blur(); // Save
            // Move selection right
            excelState.c = Math.min(excelState.c + 1, projects[activeProjectIdx][side==='A'?'dataA':'dataB'].headers.length - 1);
            renderPreviewTables();
        }
        return;
    }

    if (excelState.side !== side) return;

    let r = excelState.r;
    let c = excelState.c;
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    const maxR = data.body.length - 1;
    const maxC = data.headers.length - 1;

    // --- SHORTCUTS (When NOT editing) ---
    
    // 1. DELETE Key: Clear content
    if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        
        // Scenario A: Clear Whole Column
        if (excelState.mode === 'col') {
             if(confirm("Clear entire column?")) {
                 for(let i=0; i<=maxR; i++) data.body[i][c] = "";
                 showToast("Column Cleared");
             }
        } 
        // Scenario B: Clear Single Cell
        else if (excelState.mode === 'cell') {
            data.body[r][c] = "";
        }
        renderPreviewTables();
        return;
    }

    // 2. NAVIGATION
    if (e.key === 'ArrowUp') { r--; e.preventDefault(); }
    else if (e.key === 'ArrowDown') { r++; e.preventDefault(); }
    else if (e.key === 'ArrowLeft') { c--; e.preventDefault(); }
    else if (e.key === 'ArrowRight') { c++; e.preventDefault(); }
    else if (e.key === 'Tab') { c++; e.preventDefault(); } // Tab moves right
    
    // 3. ENTER: Start Editing
    else if (e.key === 'Enter') { 
        e.preventDefault();
        const cell = document.querySelector(`#gridContainer-${side} .excel-focus`);
        if(cell) editExcelCell(side, r, c, cell);
        return;
    }
    
    // 4. TYPE TO OVERWRITE
    // Checks for printable characters (length 1) and ensures no special keys (Ctrl/Alt) are pressed
    else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        const cell = document.querySelector(`#gridContainer-${side} .excel-focus`);
        if(cell) {
            editExcelCell(side, r, c, cell, e.key);
            e.preventDefault(); // Prevent the key from scrolling or doing other things
            return;
        }
    }

    // Update Selection bounds
    if (r < 0) r = 0; if (r > maxR) r = maxR;
    if (c < 0) c = 0; if (c > maxC) c = maxC;

    // Only re-render if selection actually changed
    if (r !== excelState.r || c !== excelState.c) {
        excelState.r = r;
        excelState.c = c;
        renderPreviewTables();
    }
}

// 6. SMART PASTE (FIXED ROW 1 BUG)
function handleExcelPaste(e, side) {
    e.preventDefault();
    e.stopPropagation();

    // 1. Force Exit Editing Mode (Crucial for fixing Row 1 bug)
    if (excelState.editing) {
        document.activeElement.blur(); 
        excelState.editing = false;
    }

    // 2. Get Data
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    const rows = text.split(/\r\n|\n|\r/).filter(r => r.trim());
    if (rows.length === 0) return;

    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;

    // SCENARIO A: Column Selected -> Pattern Fill
    if (excelState.mode === 'col') {
        const c = excelState.c;
        for (let i = 0; i < data.body.length; i++) {
            const val = rows[i % rows.length].split('\t')[0];
            data.body[i][c] = val;
        }
        showToast(`Auto-Filled ${data.body.length} rows!`);
    }
    
    // SCENARIO B: Single Cell Selected -> Normal Paste
    else if (excelState.mode === 'cell') {
        const startR = excelState.r;
        const startC = excelState.c;
        
        rows.forEach((rowStr, rOffset) => {
            const cells = rowStr.split('\t');
            cells.forEach((val, cOffset) => {
                const targetR = startR + rOffset;
                const targetC = startC + cOffset;
                
                // Update Model
                if (data.body[targetR] && targetC < data.headers.length) {
                    data.body[targetR][targetC] = val.trim();
                }
            });
        });
        showToast(`Pasted ${rows.length} rows.`);
    }
    
    renderPreviewTables();
}

function editHeader(side, colIdx, th) {
    th.contentEditable = true;
    th.focus();
    th.onblur = function() {
        const p = projects[activeProjectIdx];
        const data = side === 'A' ? p.dataA : p.dataB;
        data.headers[colIdx] = this.innerText.trim();
        th.contentEditable = false;
    };
    th.onkeydown = (e) => { if(e.key==='Enter') th.blur(); };
}

// ==========================================
// UI: SIDEBAR TOGGLE
// ==========================================
function toggleSidebar() {
    const sidebar = document.querySelector('.sidebar');
    const icon = document.querySelector('#sidebarToggle i');
    
    // Toggle the CSS class that hides the sidebar
    sidebar.classList.toggle('collapsed');
    
    // Change the icon from "hamburger" menu to "arrow" when collapsed
    if (sidebar.classList.contains('collapsed')) {
        icon.className = 'fas fa-chevron-right';
    } else {
        icon.className = 'fas fa-bars';
    }
}

// ==========================================
// UNIFIED MASTER EXCEL ENGINE
// Handles Selection, Auto-Scroll, Copy/Paste, Undo/Redo, & Fill-Handle
// ==========================================
const GridEngine = {
    activeSide: null,    
    selStart: null,      
    selEnd: null,        
    isDragging: false,
    editingCell: null,
    
    isDraggingFill: false,
    fillEndR: null,

    autoScrollTimer: null,
    mouseX: 0, mouseY: 0,
    globalEventsAttached: false,

    history: { A: { undo: [], redo: [] }, B: { undo: [], redo: [] } },

    init: function(side) {
        const container = document.getElementById(`gridContainer-${side}`);
        if (!container) return;
        
        const table = container.querySelector('.clean-table'); 
        const scrollable = container.closest('.scroll-wrap'); 
        const fillHandle = document.getElementById(`fillHandle${side}`);
        
        // FIX 2: Allow text selection ONLY when actively typing inside a cell
        table.onselectstart = (e) => {
            if (e.target.closest('.excel-editing')) return true; 
            return false; 
        };

        if (!this.globalEventsAttached) {
            document.addEventListener('mousemove', (e) => this.handleGlobalMouseMove(e));
            document.addEventListener('mouseup', (e) => this.handleGlobalMouseUp(e));
            this.globalEventsAttached = true;
        }

        container.addEventListener('mousedown', (e) => {
            if (this.editingCell || e.button !== 0) return; 
            this.activeSide = side;

            if (e.target === fillHandle) {
                this.isDraggingFill = true;
                this.fillEndR = this.selEnd.r;
                this.startAutoScroll(scrollable, table, fillHandle, side);
                return;
            }

            const cell = e.target.closest('td, th');
            if (!cell) return;

            if (cell.tagName === 'TH' && cell.hasAttribute('data-c')) {
                const c = parseInt(cell.getAttribute('data-c'));
                const maxR = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'].body.length - 1;
                this.updateSelection({ r: 0, c: c }, { r: maxR, c: c }, table, fillHandle);
                return;
            }

            if (cell.classList.contains('row-num')) {
                const r = parseInt(cell.parentElement.getAttribute('data-r'));
                const maxC = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'].headers.length - 1;
                this.updateSelection({ r: r, c: 0 }, { r: r, c: maxC }, table, fillHandle);
                return;
            }

            if (cell.hasAttribute('data-r') && cell.hasAttribute('data-c')) {
                this.isDragging = true;
                const r = parseInt(cell.getAttribute('data-r'));
                const c = parseInt(cell.getAttribute('data-c'));
                
                if (e.shiftKey && this.selStart) this.updateSelection(this.selStart, { r, c }, table, fillHandle);
                else this.updateSelection({ r, c }, { r, c }, table, fillHandle);
                
                // FIX 1: Prevent the browser from jerking the scrollbar to the left!
                container.focus({ preventScroll: true }); 
                this.startAutoScroll(scrollable, table, fillHandle, side);
            }
        });

        table.addEventListener('dblclick', (e) => {
            const cell = e.target.closest('td');
            if (cell && cell.hasAttribute('data-r')) this.startEditing(cell);
        });
        
        container.addEventListener('keydown', (e) => { this.handleKeyDown(e, table, side, fillHandle); });
        container.addEventListener('copy', (e) => { this.handleCopy(e, table); });
        container.addEventListener('paste', (e) => { this.handlePaste(e, table, side); });
    },

    handleGlobalMouseMove: function(e) {
        this.mouseX = e.clientX;
        this.mouseY = e.clientY;

        if (!this.isDragging && !this.isDraggingFill) return;
        if (!this.activeSide) return;

        const container = document.getElementById(`gridContainer-${this.activeSide}`);
        if (!container) return;
        const table = container.querySelector('.clean-table');
        const fillHandle = document.getElementById(`fillHandle${this.activeSide}`);

        const el = document.elementFromPoint(e.clientX, e.clientY);
        const cell = el ? el.closest('td') : null;

        if (cell && cell.hasAttribute('data-r') && cell.closest('table') === table) {
            const r = parseInt(cell.getAttribute('data-r'));
            const c = parseInt(cell.getAttribute('data-c'));

            if (this.isDraggingFill) {
                this.fillEndR = r;
                this.drawFillPreview(table);
            } else {
                this.updateSelection(this.selStart, { r, c }, table, fillHandle);
            }
        }
    },

    handleGlobalMouseUp: function(e) {
        if (this.isDraggingFill && this.activeSide) this.executeFill(this.activeSide);
        this.isDragging = false; 
        this.isDraggingFill = false;
        this.stopAutoScroll();
    },

    startAutoScroll: function(scrollable, table, fillHandle, side) {
        this.stopAutoScroll();
        
        this.autoScrollTimer = setInterval(() => {
            if (!this.isDragging && !this.isDraggingFill) return;

            const rect = scrollable.getBoundingClientRect();
            const buffer = 50; 
            const speed = 25;  
            let scrolled = false;

            if (this.mouseY > rect.bottom - buffer) { scrollable.scrollTop += speed; scrolled = true; }
            else if (this.mouseY < rect.top + buffer) { scrollable.scrollTop -= speed; scrolled = true; }

            if (this.mouseX > rect.right - buffer) { scrollable.scrollLeft += speed; scrolled = true; }
            else if (this.mouseX < rect.left + buffer) { scrollable.scrollLeft -= speed; scrolled = true; }

            const clampX = Math.max(rect.left + 5, Math.min(this.mouseX, rect.right - 25));
            const clampY = Math.max(rect.top + 5, Math.min(this.mouseY, rect.bottom - 25));

            if (fillHandle) fillHandle.style.pointerEvents = 'none';
            const el = document.elementFromPoint(clampX, clampY);
            if (fillHandle) fillHandle.style.pointerEvents = 'auto';

            const cell = el ? el.closest('td') : null;

            if (cell && cell.hasAttribute('data-r') && cell.closest('table') === table) {
                const r = parseInt(cell.getAttribute('data-r'));
                const c = parseInt(cell.getAttribute('data-c'));
                if (this.isDraggingFill) {
                    this.fillEndR = r;
                    this.drawFillPreview(table);
                } else {
                    this.updateSelection(this.selStart, { r, c }, table, fillHandle);
                }
            }
        }, 30);
    },

    stopAutoScroll: function() {
        if (this.autoScrollTimer) clearInterval(this.autoScrollTimer);
        this.autoScrollTimer = null;
    },

    updateSelection: function(start, end, table, fillHandle) {
        this.selStart = start; this.selEnd = end;
        
        table.querySelectorAll('.excel-selected, .excel-focus, .excel-border-top, .excel-border-bottom, .excel-border-left, .excel-border-right, .excel-fill-preview')
             .forEach(el => el.classList.remove('excel-selected', 'excel-focus', 'excel-border-top', 'excel-border-bottom', 'excel-border-left', 'excel-border-right', 'excel-fill-preview'));

        const rMin = Math.min(this.selStart.r, this.selEnd.r);
        const rMax = Math.max(this.selStart.r, this.selEnd.r);
        const cMin = Math.min(this.selStart.c, this.selEnd.c);
        const cMax = Math.max(this.selStart.c, this.selEnd.c);

        let brCell = null;

        for (let r = rMin; r <= rMax; r++) {
            for (let c = cMin; c <= cMax; c++) {
                const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                if (!cell) continue;

                cell.classList.add('excel-selected');
                if (r === rMin) cell.classList.add('excel-border-top');
                if (r === rMax) cell.classList.add('excel-border-bottom');
                if (c === cMin) cell.classList.add('excel-border-left');
                if (c === cMax) cell.classList.add('excel-border-right');
                if (r === this.selEnd.r && c === this.selEnd.c) cell.classList.add('excel-focus');
                
                if (r === rMax && c === cMax) brCell = cell;
            }
        }

        if (brCell && fillHandle) {
            fillHandle.style.display = 'block';
            fillHandle.style.top = (brCell.offsetTop + brCell.offsetHeight - 5) + 'px';
            fillHandle.style.left = (brCell.offsetLeft + brCell.offsetWidth - 5) + 'px';
        }
    },

    handleKeyDown: function(e, table, side, fillHandle) {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); this.undo(side); return; }
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); this.redo(side); return; }

        if (this.editingCell) {
            if (e.key === 'Enter') { e.preventDefault(); this.stopEditing(true, side); }
            else if (e.key === 'Escape') { e.preventDefault(); this.stopEditing(false, side); }
            return;
        }

        if (!this.selEnd) return;

        let r = this.selEnd.r; let c = this.selEnd.c;
        const maxR = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'].body.length - 1;
        const maxC = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'].headers.length - 1;

        if (e.key === 'Delete' || e.key === 'Backspace') {
            e.preventDefault();
            this.clearSelection(side);
            return;
        }

        if (e.key.startsWith('Arrow') || e.key === 'Tab') {
            e.preventDefault();
            if (e.key === 'ArrowUp') r--;
            if (e.key === 'ArrowDown') r++;
            if (e.key === 'ArrowLeft' || (e.key === 'Tab' && e.shiftKey)) c--;
            if (e.key === 'ArrowRight' || (e.key === 'Tab' && !e.shiftKey)) c++;

            if (r < 0) r = 0; if (r > maxR) r = maxR;
            if (c < 0) c = 0; if (c > maxC) c = maxC;

            if (e.shiftKey && e.key.startsWith('Arrow')) this.updateSelection(this.selStart, { r, c }, table, fillHandle);
            else this.updateSelection({ r, c }, { r, c }, table, fillHandle);
            
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell) cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }
        else if (e.key === 'Enter') {
            e.preventDefault();
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell) this.startEditing(cell);
        }
        else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            e.preventDefault();
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell) {
                this.startEditing(cell);
                cell.innerText = e.key; 
                const sel = window.getSelection();
                const range = document.createRange();
                range.selectNodeContents(cell);
                range.collapse(false);
                sel.removeAllRanges();
                sel.addRange(range);
            }
        }
    },

    handleCopy: function(e, table) {
        if (this.editingCell || !this.selStart || !this.selEnd) return;
        e.preventDefault();

        const rMin = Math.min(this.selStart.r, this.selEnd.r);
        const rMax = Math.max(this.selStart.r, this.selEnd.r);
        const cMin = Math.min(this.selStart.c, this.selEnd.c);
        const cMax = Math.max(this.selStart.c, this.selEnd.c);

        let tsv = "";
        for (let r = rMin; r <= rMax; r++) {
            let rowText = [];
            for (let c = cMin; c <= cMax; c++) {
                const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                rowText.push(cell ? cell.innerText.trim() : "");
            }
            tsv += rowText.join("\t") + "\n";
        }
        e.clipboardData.setData('text/plain', tsv.trimEnd());
        showToast(`Copied ${rMax - rMin + 1} rows!`);
    },

    handlePaste: function(e, table, side) {
        if (this.editingCell || !this.selEnd) return;
        e.preventDefault();

        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (!text) return;

        this.saveState(side);

        const pasteRows = text.split(/\r\n|\n|\r/).filter(r => r.trim() !== "");
        const startR = Math.min(this.selStart.r, this.selEnd.r);
        const startC = Math.min(this.selStart.c, this.selEnd.c);
        
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
        let rowsUpdated = 0;

        pasteRows.forEach((rowStr, rOffset) => {
            const targetR = startR + rOffset;
            if (targetR >= data.body.length) return; 
            
            const pasteCols = rowStr.split('\t');
            pasteCols.forEach((val, cOffset) => {
                const targetC = startC + cOffset;
                if (targetC < data.headers.length) data.body[targetR][targetC] = val.trim();
            });
            rowsUpdated++;
        });

        renderPreviewTables();
        showToast(`Pasted ${rowsUpdated} rows!`);
    },

    clearSelection: function(side) {
        if (!this.selStart || !this.selEnd) return;
        this.saveState(side);

        const rMin = Math.min(this.selStart.r, this.selEnd.r);
        const rMax = Math.max(this.selStart.r, this.selEnd.r);
        const cMin = Math.min(this.selStart.c, this.selEnd.c);
        const cMax = Math.max(this.selStart.c, this.selEnd.c);

        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];

        for (let r = rMin; r <= rMax; r++) {
            for (let c = cMin; c <= cMax; c++) {
                if (data.body[r] && c < data.headers.length) data.body[r][c] = "";
            }
        }
        renderPreviewTables();
        showToast("Cleared Selection");
    },

    drawFillPreview: function(table) {
        table.querySelectorAll('.excel-fill-preview').forEach(el => el.classList.remove('excel-fill-preview'));
        if (this.fillEndR === null || !this.selEnd) return;

        const srcMaxR = Math.max(this.selStart.r, this.selEnd.r);
        const targetRMax = Math.max(srcMaxR, this.fillEndR);
        const cMin = Math.min(this.selStart.c, this.selEnd.c);
        const cMax = Math.max(this.selStart.c, this.selEnd.c);

        for (let r = srcMaxR + 1; r <= targetRMax; r++) {
            for (let c = cMin; c <= cMax; c++) {
                const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                if (cell) cell.classList.add('excel-fill-preview');
            }
        }
    },

    executeFill: function(side) {
        if (this.fillEndR === null || !this.selEnd) return;
        
        const srcMinR = Math.min(this.selStart.r, this.selEnd.r);
        const srcMaxR = Math.max(this.selStart.r, this.selEnd.r);
        const cMin = Math.min(this.selStart.c, this.selEnd.c);
        const cMax = Math.max(this.selStart.c, this.selEnd.c);
        const targetRMax = Math.max(srcMaxR, this.fillEndR);

        if (targetRMax <= srcMaxR) return;

        this.saveState(side);
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
        const srcRowCount = (srcMaxR - srcMinR) + 1;

        for (let r = srcMaxR + 1; r <= targetRMax; r++) {
            const patternR = srcMinR + ((r - srcMaxR - 1) % srcRowCount);
            for (let c = cMin; c <= cMax; c++) {
                if (data.body[r] && data.body[patternR]) {
                    data.body[r][c] = data.body[patternR][c];
                }
            }
        }
        
        this.selEnd.r = targetRMax;
        renderPreviewTables();
        showToast("Auto-Filled Data");
    },

    saveState: function(side) {
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
        const snapshot = data.body.map(row => {
            const newRow = [...row];
            Object.defineProperty(newRow, '_originalIdx', { value: row._originalIdx, writable: true, enumerable: false });
            return newRow;
        });

        this.history[side].undo.push(snapshot);
        if (this.history[side].undo.length > 20) this.history[side].undo.shift(); 
        this.history[side].redo = []; 
        document.getElementById(`undoBadge${side}`).style.display = 'inline';
    },

    undo: function(side) {
        if (this.history[side].undo.length === 0) return showToast("Nothing to undo");
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
        
        const currentSnapshot = data.body.map(r => { const nr = [...r]; Object.defineProperty(nr, '_originalIdx', { value: r._originalIdx, writable: true, enumerable: false }); return nr; });
        this.history[side].redo.push(currentSnapshot);

        data.body = this.history[side].undo.pop();
        if (this.history[side].undo.length === 0) document.getElementById(`undoBadge${side}`).style.display = 'none';
        renderPreviewTables();
        showToast("Undid last action");
    },

    redo: function(side) {
        if (this.history[side].redo.length === 0) return showToast("Nothing to redo");
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
        
        const currentSnapshot = data.body.map(r => { const nr = [...r]; Object.defineProperty(nr, '_originalIdx', { value: r._originalIdx, writable: true, enumerable: false }); return nr; });
        this.history[side].undo.push(currentSnapshot);

        data.body = this.history[side].redo.pop();
        document.getElementById(`undoBadge${side}`).style.display = 'inline';
        renderPreviewTables();
        showToast("Redid action");
    },

    startEditing: function(cell) {
        this.editingCell = cell;
        cell.contentEditable = true;
        cell.classList.add('excel-editing');
        cell.focus();
        cell.setAttribute('data-original', cell.innerText); 
        cell.onblur = () => { if (this.editingCell === cell) this.stopEditing(true, this.activeSide); };
    },

    stopEditing: function(saveValue, side) {
        if (!this.editingCell) return;
        const cell = this.editingCell;
        cell.contentEditable = false;
        cell.classList.remove('excel-editing');
        cell.onblur = null;

        if (!saveValue) {
            cell.innerText = cell.getAttribute('data-original');
        } else {
            if (cell.innerText !== cell.getAttribute('data-original')) this.saveState(side); 
            
            const r = parseInt(cell.getAttribute('data-r'));
            const c = parseInt(cell.getAttribute('data-c'));
            const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
            
            if (data && data.body[r]) data.body[r][c] = cell.innerText.replace(/[\r\n]+/g, " ").trim();
        }
        
        this.editingCell = null;
        const container = document.getElementById(`gridContainer-${side}`);
        // FIX 1: Prevent jump when finishing edit
        if(container) container.focus({ preventScroll: true }); 
    }
};

function getVisibleExcelData(sheet) {
    let data = extractSmartExcelData(sheet);

    // 1. Remove Hidden Rows
    if (sheet['!rows'] && data.length > 0) {
        data = data.filter((row, rowIndex) => {
            const rMeta = sheet['!rows'][rowIndex];
            const isHidden = rMeta && (rMeta.hidden === true || rMeta.hidden === 1 || rMeta.hpx === 0 || rMeta.ht === 0);
            return !isHidden;
        });
    }

    // 2. Remove Hidden Columns
    if (sheet['!cols'] && data.length > 0) {
        const hiddenColIndices = new Set();
        sheet['!cols'].forEach((col, i) => {
            if (col && (col.hidden === true || col.hidden === 1 || col.wpx === 0 || col.width === 0)) {
                hiddenColIndices.add(i);
            }
        });
        if (hiddenColIndices.size > 0) {
            data = data.map(row => row.filter((_, i) => !hiddenColIndices.has(i)));
        }
    }

    return data;
}