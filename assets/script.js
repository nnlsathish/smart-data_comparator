// ==========================================
// GLOBAL STATE & APP INITIALIZATION
// ==========================================

// Every comparison "Set" (a Form vs CPQ pair) lives in `projects`; `activeProjectIdx`
// points at the one currently on screen.
let projects = [];
let activeProjectIdx = 0;
let isOverviewMode = false;
const MAX_TOTAL_SETS = 100;

// The workbook/sheet currently open in the manual table-selection modal.
let manualWorkbook = null;
let manualFilename = "";
let currentSheetName = "";
let selectedHeaderRowIndex = -1;

// Manual-selection queue (for batches where several files fail auto-detection).
let manualQueue = [];        // [{ workbook, filename }] files still awaiting manual selection
let manualQueueTotal = 0;    // how many files needed manual selection this batch
let manualBatchActive = false;
let batchStartCount = 0;     // projects.length when the current upload batch began

// Drag-to-select state for the manual table selector.
let isDragging = false;
let selectionStartRow = -1;
let selectionEndRow = -1;

// State backing the "Undo Delete" feature (last removed row and row history).
let deletedItem = null;
let deletedItemIdx = -1;
let deletedRowsHistory = [];


// Last uploaded workbook, kept so manual selection can fall back to it.
let lastUploadedWorkbook = null;
let lastUploadedFilename = "";
let editingProjectIndex = -1;

// App entry point: create the first empty Set and wire up page-level listeners.
window.onload = function () {
    console.log("Validation Tool Loaded Successfully.");

    // Start with one empty project so the UI has something to show.
    addNewProject();

    // Warn before leaving the page if there is unsaved work.
    window.addEventListener('beforeunload', function (e) {
        if (hasUnsavedData()) {
            e.preventDefault();
            e.returnValue = 'Unsaved changes';
            return 'Unsaved changes';
        }
    });

    // Pressing Enter while the custom modal is open confirms it.
    window.addEventListener('keydown', function (event) {
        const modal = document.getElementById('customModal');
        if (event.key === 'Enter' && modal && modal.classList.contains('open')) {
            event.preventDefault();
            document.getElementById('modalBtn').click();
        }
    });
};

// True if there is any data (in the textareas or in a saved Set) worth warning about.
function hasUnsavedData() {
    const inputA = document.getElementById('tableA');
    const inputB = document.getElementById('tableB');

    const hasVisibleData = (inputA?.value.trim().length > 0) || (inputB?.value.trim().length > 0);
    const hasStoredData = projects.some(p => p.status !== 'empty');

    return hasVisibleData || hasStoredData;
}

// ==========================================
// SET MANAGEMENT (reset, create, add)
// ==========================================

// "Start Over": wipe all Sets, inputs and UI state, then create a fresh empty Set.
function hardResetApp() {
    if (!confirm("Start Fresh? This will delete all current sets and data.")) {
        return;
    }

    // Reset all in-memory state.
    projects = [];
    activeProjectIdx = 0;
    isOverviewMode = false;
    deletedItem = null;
    deletedItemIdx = -1;
    deletedRowsHistory = [];

    // Clear the hidden file inputs so re-uploading the same file still fires 'change'.
    document.querySelectorAll('input[type="file"]').forEach(el => el.value = "");

    // Hide the "Restore Set" button.
    const restoreBtn = document.getElementById('btnRestoreSet');
    if (restoreBtn) {
        restoreBtn.style.display = 'none';
        restoreBtn.innerHTML = '';
    }

    // Clear the custom quantity-keyword input.
    const customKeyInput = document.getElementById('customQtyKey');
    if (customKeyInput) customKeyInput.value = "";

    // Reset sidebar/overview UI back to the single-Set view.
    const sidebar = document.querySelector('.sidebar');
    if (sidebar) sidebar.classList.remove('overview-locked');

    const tabOverview = document.getElementById('tabOverview');
    if (tabOverview) tabOverview.classList.remove('active');

    const overviewSection = document.getElementById('overviewSection');
    if (overviewSection) overviewSection.style.display = 'none';

    // Show Step 1 and hide every other step section.
    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    const step1 = document.getElementById('step1');
    if (step1) step1.style.display = 'block';

    // Highlight Step 1 in the sidebar nav.
    document.querySelectorAll('.v-step').forEach(el => el.classList.remove('active'));
    const navStep1 = document.getElementById('navStep1');
    if (navStep1) navStep1.classList.add('active');

    // Create the first Set of the new session.
    createSet();
    renderTopBar();

    // Show the fresh empty project (also clears the textareas and tables).
    loadProjectIntoView(0);

    showToast("Application Reset Successfully");
}

// Keeps the "+ Add N Sets" button label in sync with the quantity input.

function updateAddButtonText() {
    const input = document.getElementById('addSetQty');
    const btn = document.getElementById('btnAddSets');
    if (input && btn) {
        let qty = parseInt(input.value) || 0;
        if (qty < 1) qty = 1;
        btn.innerText = `+ Add ${qty} Set${qty > 1 ? 's' : ''}`;
    }
}

// Sets the header-parsing mode ("1row"/"2row"/etc.) for side A or B and
// highlights the matching toggle button.
function setMode(side, mode) {
    document.getElementById(`headerMode${side}`).value = mode;
    const group = document.getElementById(`groupMode${side}`);
    group.querySelectorAll('.toggle-btn').forEach(btn => {
        if (btn.getAttribute('data-val') === mode) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

// Creates however many Sets the quantity input requests, then jumps to the first new one.
function addSetsFromInput() {
    const input = document.getElementById('addSetQty');
    let qty = parseInt(input.value) || 1;

    if (qty < 1) qty = 1;

    if (projects.length + qty > MAX_TOTAL_SETS) {
        showModal("Limit Reached", `Cannot add more than ${MAX_TOTAL_SETS} sets.`, 'error');
        return;
    }

    for (let i = 0; i < qty; i++) {
        createSet();
    }

    renderTopBar();
    switchProject(projects.length - qty);
    input.value = 1;
    updateAddButtonText();
}

// Appends one empty Set and switches to it.
function addNewProject() {
    if (projects.length >= MAX_TOTAL_SETS) return;
    createSet();
    renderTopBar();
    switchProject(projects.length - 1);
}


// Pushes a new project object onto `projects`. Each Set carries its own raw text,
// parsed data, mapping, matrix rules and independent parsing settings.
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
        // Per-Set settings, so one Set's mode/trim choices don't affect another.
        settings: {
            modeA: '1row',
            modeB: '1row',
            trimResults: false,
            showErrors: false
        }
    });
}

// ==========================================
// FILE UPLOAD (Form / Table 1 side)
// ==========================================

// Reads one or more uploaded Excel files, auto-detects the data table in each sheet,
// and creates a Set per detected table. Files whose table can't be found are queued
// for manual selection.
function handleBulkUpload(input) {
    if (!input.files || input.files.length === 0) return;

    // Persist whatever is on screen before we start mutating projects.
    saveCurrentViewToProject();

    // If the only Set is the initial empty one, drop it so uploads start clean.
    const currentA = document.getElementById('tableA').value.trim();
    const currentB = document.getElementById('tableB').value.trim();
    if (projects.length === 1 && projects[0].status === 'empty' && !projects[0].rawA && !currentA && !currentB) {
        projects = [];
        activeProjectIdx = -1;
    }

    // Optional user-supplied keyword used to locate the quantity column.
    const customKeyInput = document.getElementById('customQtyKey');
    const customKey = customKeyInput ? customKeyInput.value.trim() : "";

    // Process files oldest-first so multi-file batches land in download order.
    const files = Array.from(input.files).sort((a, b) => a.lastModified - b.lastModified);

    let totalCreated = 0;

    // Remembered for the manual-selection fallback.
    let tempWorkbook = null;
    let tempFilename = "";

    // Fresh manual-selection queue for this upload batch
    manualQueue = [];
    manualQueueTotal = 0;
    manualBatchActive = false;
    batchStartCount = projects.length; // baseline so we can count sets created this batch

    // Reads files one at a time (recursively) so FileReader stays sequential.
    const processNextFile = (fileIdx) => {
        if (fileIdx >= files.length) {
            input.value = ""; // allow re-selecting the same file later

            // Keep the last workbook for the manual-selection fallback.
            if (tempWorkbook) {
                lastUploadedWorkbook = tempWorkbook;
                lastUploadedFilename = tempFilename;
            }

            // All files read — decide what to show next.
            renderTopBar();

            // If CPQ files were queued for a folder batch, inject them now
            // (only when nothing still needs manual selection).
            if (totalCreated > 0 && manualQueue.length === 0 &&
                window.pendingBatchCPQs && window.pendingBatchCPQs.length > 0) {
                processBatchCPQs(window.pendingBatchCPQs, totalCreated);
                window.pendingBatchCPQs = null;
                return; // skip the standard summary modal
            }

            // Focus the first auto-created Set, if any.
            if (totalCreated > 0) {
                switchProject(projects.length - totalCreated, true);
            }

            // Walk the user through every file that failed auto-detection, one by one.
            // (Manual selection creates those Sets, so no empty placeholder is needed.)
            if (manualQueue.length > 0) {
                startManualQueue();
                return;
            }

            // Nothing imported and nothing to select — keep one empty Set around.
            if (totalCreated === 0 && projects.length === 0) {
                createSet(); activeProjectIdx = 0;
            }

            // Otherwise show the import summary.
            if (totalCreated > 0) {
                const msg = `✅ Successfully loaded <strong>${totalCreated}</strong> sheets from <strong>${files.length}</strong> file(s).`;
                showModal("Import Complete", msg, "success");
            }
            return;
        }

        const file = files[fileIdx];
        tempFilename = file.name;
        const reader = new FileReader();

        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                tempWorkbook = workbook;

                // Remember how many sets existed before this file, so we can tell
                // whether THIS file produced anything (if not, it needs manual selection)
                const createdBeforeThisFile = totalCreated;

                // Optionally sort sheets A–Z (numeric-aware) within the file.
                let sheetNamesToProcess = [...workbook.SheetNames];
                const sortToggle = document.getElementById('sortSheetsToggle');

                if (sortToggle && sortToggle.checked) {
                    sheetNamesToProcess.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
                }

                sheetNamesToProcess.forEach((sheetName) => {
                    // Ignore sheets Excel marks as hidden.
                    if (workbook.Workbook && workbook.Workbook.Sheets) {
                        const sMeta = workbook.Workbook.Sheets.find(s => s.name === sheetName);
                        if (sMeta && (sMeta.Hidden !== 0 || sMeta.state === 'hidden')) return;
                    }

                    const sheet = workbook.Sheets[sheetName];

                    // Extract rows, preserving barcodes (avoids scientific-notation mangling).
                    let rawData = typeof extractSmartExcelData === "function"
                        ? extractSmartExcelData(sheet)
                        : XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

                    // Drop rows Excel hid (explicit hidden flag or zero height).
                    if (sheet['!rows'] && rawData.length > 0) {
                        rawData = rawData.filter((row, rowIndex) => {
                            const rMeta = sheet['!rows'][rowIndex];
                            const isHidden = rMeta && (rMeta.hidden === true || rMeta.hidden === 1 || rMeta.hpx === 0 || rMeta.ht === 0);
                            return !isHidden;
                        });
                    }

                    // Drop hidden columns, but keep the feature-block columns between
                    // "Feature Field 1" and "Total Qty" even when they're hidden.
                    if (sheet['!cols'] && rawData.length > 0) {
                        const hiddenIndices = new Set();
                        let startIdx = -1;
                        let endIdx = -1;

                        for (let r = 0; r < Math.min(rawData.length, 30); r++) {
                            const row = rawData[r];
                            if (!row || !Array.isArray(row)) continue;

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

                    // Locate and extract just the data table from the raw sheet.
                    const cleanTable = extractTableStrict(rawData, customKey);

                    // If auto-split is on, a single sheet may hold several tables; split
                    // them into separate Sets and skip the normal single-table path.
                    const autoSplitToggle = document.getElementById('chkAutoSplit');
                    if (autoSplitToggle && autoSplitToggle.checked) {
                        let splitCount = performSmartAutoSplit(rawData, cleanTable, file.name, sheetName, workbook);
                        if (splitCount > 0) {
                            totalCreated += splitCount;
                            return; // this sheet's tables are handled by the splitter
                        }
                    }

                    // Treat the sheet as valid only if its quantity column has a non-zero value.
                    let hasValidData = false;
                    if (cleanTable.length > 1) {
                        const headerRow = cleanTable[0];
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

                        // Disambiguate only when the same file has two identically named sheets.
                        let finalName = sheetName;
                        let counter = 1;
                        while (projects.some(p => p.name === finalName && p.fileName === file.name)) {
                            finalName = `${sheetName} (${counter})`;
                            counter++;
                        }

                        // Create the Set, passing the workbook so it remembers its source file.
                        createSet(sheetName, workbook, sheetName, file.name);
                        let p = projects[projects.length - 1];

                        // Tag with the upload session so the top bar groups this batch together.
                        p.uploadSessionId = input.uploadSessionId || Date.now();

                        p.rawA = arrayToTSV(cleanTable);

                        // Pull any "key: value" matrix rules from above the table header.
                        try {
                            const headerIdx = findHeaderRowIndex(rawData);
                            if (headerIdx > 0) {
                                const matrixData = extractMatrixData(rawData, headerIdx);
                                if (matrixData.length > 0) {
                                    p.rawMatrix = matrixData.map(m => `${m.key}: ${m.val}`).join("\n");
                                    p.matrix = parseMatrixString(p.rawMatrix);
                                    p.showMatrix = true;
                                }
                            }
                        } catch (e) { }

                        p.status = 'ready';
                    }
                });

                // File produced no tables across all sheets -> queue for manual selection.
                if (totalCreated === createdBeforeThisFile) {
                    manualQueue.push({ workbook: workbook, filename: file.name });
                }

            } catch (err) {
                console.error("Error reading file:", file.name, err);
            } finally {
                processNextFile(fileIdx + 1);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    processNextFile(0); // start the sequential read
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

    // Scan every visible sheet for a table containing the keyword column.
    workbook.SheetNames.forEach((sheetName) => {
        if (workbook.Workbook && workbook.Workbook.Sheets) {
            const sMeta = workbook.Workbook.Sheets.find(s => s.name === sheetName);
            if (sMeta && (sMeta.Hidden !== 0)) return;
        }

        const sheet = workbook.Sheets[sheetName];
        let rawData = extractSmartExcelData(sheet);
        const cleanTable = extractTableStrict(rawData, customKey);

        // A real table has more than just a header row.
        if (cleanTable.length > 1) {
            matchedSheets.push({ name: sheetName, data: cleanTable });
        }
    });

    if (matchedSheets.length === 0) {
        showModal("Not Found", `No sheets found containing column: <strong>${customKey}</strong>`, "error");
    } else {
        showSheetSelector(matchedSheets, customKey);
    }
}

// Reuses the manual-selection modal to let the user pick which keyword-matched
// sheets to import, each as its own Set.
function showSheetSelector(matches, keyword) {
    const modal = document.getElementById('manualSelectModal');
    const tableArea = document.getElementById('manualRawTable');
    const titleArea = modal.querySelector('h3');
    const instruction = document.getElementById('manualInstruction');
    const confirmBtn = document.getElementById('btnManualConfirm');

    // Set up the modal header/instructions.
    modal.style.display = 'flex';
    titleArea.innerHTML = `<i class="fas fa-search"></i> Found "${keyword}" in ${matches.length} Sheets`;
    instruction.innerHTML = `<span style="color:#2563eb">Select sheets to import as separate sets:</span>`;

    // Build the checkbox list of matched sheets.
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

    // Swap the raw-table view out for our checkbox list.
    const container = tableArea.parentElement;
    tableArea.style.display = 'none';

    // Remove any list from a previous open so we start fresh.
    const existingList = document.getElementById('customSheetList');
    if (existingList) existingList.remove();

    const listDiv = document.createElement('div');
    listDiv.id = 'customSheetList';
    listDiv.style.flex = "1";
    listDiv.style.overflow = "hidden";
    listDiv.innerHTML = html;
    container.appendChild(listDiv);

    // Turn the confirm button into "Import Selected" for this modal mode.
    confirmBtn.disabled = false;
    confirmBtn.innerHTML = "Import Selected";
    confirmBtn.style.background = "#2563eb";
    confirmBtn.style.cursor = "pointer";

    // "Select All" checkbox behaviour.
    window.toggleAllSheets = function (source) {
        document.querySelectorAll('.sheet-chk').forEach(c => c.checked = source.checked);
    };

    // Import each checked sheet: reuse the active Set for the first, new Sets for the rest.
    confirmBtn.onclick = function () {
        const checkboxes = document.querySelectorAll('.sheet-chk:checked');
        if (checkboxes.length === 0) {
            alert("Please select at least one sheet.");
            return;
        }

        checkboxes.forEach((chk, i) => {
            const match = matches[parseInt(chk.value)];
            let targetIdx = activeProjectIdx;

            if (i > 0) {
                // Use createSet so each new tab keeps its source-file memory.
                createSet(match.name, lastUploadedWorkbook, match.name, lastUploadedFilename);
                targetIdx = projects.length - 1;
            }

            const p = projects[targetIdx];
            p.name = match.name;
            p.fileName = lastUploadedFilename;
            p.rawA = arrayToTSV(match.data);
            p.status = 'ready';
            p.step = 1;
        });

        // Restore the modal to its normal manual-selection state.
        document.getElementById('customSheetList').remove();
        tableArea.style.display = 'block';
        closeManualModal();

        loadProjectIntoView(activeProjectIdx);
        renderTopBar();
        showToast(`Imported ${checkboxes.length} sheets successfully!`);

        // Put the confirm button back to its default "Confirm & Import" action.
        confirmBtn.onclick = confirmManualImport;
        confirmBtn.innerHTML = "Confirm & Import";
    };
}

// ==========================================
// EXCEL EXTRACTION HELPERS
// ==========================================
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
// TABLE DETECTION & EXTRACTION
// ==========================================

// Locates the real data table inside a messy sheet (skipping address blocks, notes,
// etc.), merges two-row headers into single column names, and returns clean rows.
// `customKeyword`, if given, is used instead of the default quantity regex to find
// the header row.
function extractTableStrict(data, customKeyword = null) {
    if (!data || !Array.isArray(data) || data.length === 0) return [];
    data = data.filter(r => Array.isArray(r));

    let startIndex = -1;

    // Regex that identifies the header row by its quantity column.
    let qtyRegex;
    if (customKeyword) {
        const escaped = customKeyword.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const flexSpace = escaped.replace(/\s+/g, '\\s*'); 
        qtyRegex = new RegExp(flexSpace, "i");
    } else {
        qtyRegex = /qty|quantity|total\s*qty|bill\s*(?:&|and)?\s*ship\s*(?:qty|quantity)|round\s*up|order\s*(?:qty|quantity)/i;
    }
        
    const chineseRegex = /[\u4e00-\u9fff]/;

    // If a "SKU Information" marker exists, start scanning just below it.
    let searchStartRow = 0;
    for (let i = 0; i < Math.min(data.length, 30); i++) {
        if (data[i] && data[i].join(" ").toLowerCase().includes("sku information")) {
            searchStartRow = i + 1;
            break;
        }
    }

    // Find the header row: the first row with a quantity column that isn't a
    // note/address/instruction line.
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

    // Merge a two-row header into one, then slice out the data rows.

    let headerRow = data[startIndex] || [];
    let nextRow = data[startIndex + 1];
    let dataStartIndex = startIndex + 1;

    // Skip an "#N/A" filler row between the header and the data.
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

        // If the next row has a number in the quantity column it's a data row,
        // not a second header line — so don't merge it.
        const qtyIdx = headerRow.findIndex(c => c && qtyRegex.test(String(c).trim()));
        if (qtyIdx !== -1 && nextRow[qtyIdx] !== undefined && nextRow[qtyIdx] !== null) {
            const nextRowQtyVal = String(nextRow[qtyIdx]).trim().replace(/,/g, '');
            if (/^-?\d+(\.\d+)?$/.test(nextRowQtyVal)) {
                useCombinedHeader = false;
            }
        }
    }

    // Build one header name per column, combining the two header rows when needed.
    let unifiedHeaders = [];
    const maxLen = Math.max(headerRow.length || 0, (nextRow ? nextRow.length : 0));
    let lastMainHeader = "";

    for (let c = 0; c < maxLen; c++) {
        let val1 = headerRow[c] ? String(headerRow[c]).trim() : "";
        let val2 = (useCombinedHeader && nextRow && nextRow[c]) ? String(nextRow[c]).trim() : "";

        // Inherit the previous main header only for columns that have a sub-header,
        // so a blank column doesn't "bleed" the name across unrelated columns.
        if (!val1 && lastMainHeader && val2) {
            val1 = lastMainHeader;
        } else if (val1) {
            lastMainHeader = val1;
        } else {
            lastMainHeader = "";
        }

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

    // Collect the data rows, stopping at the footer/total block and skipping junk rows.
    let rawBody = [];
    let emptyRowCount = 0;

    for (let i = dataStartIndex; i < data.length; i++) {
        let row = data[i];

        if (!row || row.every(c => !c || String(c).trim() === "")) {
            emptyRowCount++;
            // Tolerate large blank gaps; give up only after 50 empty rows in a row.
            if (emptyRowCount >= 50) break;
            continue;
        }
        emptyRowCount = 0;

        const rowStr = row.join(" ").toLowerCase();

        // Stop at footer/summary lines so we don't ingest notes or totals as data.
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

        // Skip placeholder/example rows (e.g. "e.g.", "ex:", "example:").
        let isExampleRow = row.some(c => {
            if (!c) return false;
            let text = String(c).toLowerCase().trim();
            return /^(e\.?g\.?|ex\.|ex:|example\b)/i.test(text) || text.includes("e.g.") || text.includes("(eg") || text.includes("example:");
        });
        if (isExampleRow) continue;

        rawBody.push(row);
    }

    // Group adjacent columns that share the same merged header name.
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

    // For each same-named group keep only columns with unique data, dropping empty
    // duplicate clones while preserving genuinely distinct columns.
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

    // Number repeated header names so each column is unique (e.g. "Gender 1", "Gender 2").
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

    // Drop auto-named "Column X" ghost columns unless they actually hold data.
    let indicesToKeep = [];
    let namesToKeep = [];

    for (let i = 0; i < finalIndices.length; i++) {
        let colName = finalHeaderNames[i];
        let originalIdx = finalIndices[i];

        if (colName.toLowerCase().startsWith("column ")) {
            let hasData = rawBody.some(row => {
                let val = row[originalIdx];
                return val !== null && val !== undefined && String(val).trim() !== "";
            });

            if (hasData) {
                indicesToKeep.push(originalIdx);
                namesToKeep.push(colName);
            }
        } else {
            indicesToKeep.push(originalIdx);
            namesToKeep.push(colName);
        }
    }

    finalIndices = indicesToKeep;
    finalHeaderNames = namesToKeep;

    // Assemble the final [header, ...rows] table, padded to a uniform width.
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
// STEP 1 LIVE QUANTITY TRACKER
// ==========================================

// Parses the Step 1 textarea for `side`, lists its quantity columns in a dropdown
// (shown only when there are several), and updates the live total.
function scanForLiveQty(side) {
    const raw = document.getElementById(side === 'A' ? 'tableA' : 'tableB').value;
    const selectEl = document.getElementById(side === 'A' ? 'qtySelectA' : 'qtySelectB');
    const qtyDisplay = document.getElementById(side === 'A' ? 'liveQtyA' : 'liveQtyB');

    if (!raw || !raw.trim()) {
        if(selectEl) selectEl.style.display = 'none';
        if(qtyDisplay) qtyDisplay.innerText = "0";
        return;
    }

    const data = parseExcelData(raw, "1row");
    if (!data || !data.headers) return;

    // Collect every quantity-like column.
    const qtyRegex = /qty|quantity|units|pcs|order/i;
    let qtyCols = [];
    data.headers.forEach((h, i) => {
        if (h && qtyRegex.test(String(h))) qtyCols.push({name: h, index: i});
    });

    if (qtyCols.length === 0) {
        selectEl.style.display = 'none';
        qtyDisplay.innerText = "0";
        return;
    }

    // Rebuild the dropdown, preserving the user's current choice if still valid.
    const currentSel = selectEl.value;
    selectEl.innerHTML = "";

    qtyCols.forEach(c => {
        let opt = document.createElement('option');
        opt.value = c.index;
        opt.innerText = c.name;
        selectEl.appendChild(opt);
    });

    if (qtyCols.some(c => c.index.toString() === currentSel)) {
        selectEl.value = currentSel;
    }

    // Only expose the picker when there is more than one quantity column.
    selectEl.style.display = qtyCols.length > 1 ? 'inline-block' : 'none';

    calculateLiveQty(side, data);
}

function calculateLiveQty(side, parsedData = null) {
    if (!parsedData) {
        const raw = document.getElementById(side === 'A' ? 'tableA' : 'tableB').value;
        parsedData = parseExcelData(raw, "1row");
    }
    
    const selectEl = document.getElementById(side === 'A' ? 'qtySelectA' : 'qtySelectB');
    const qtyDisplay = document.getElementById(side === 'A' ? 'liveQtyA' : 'liveQtyB');
    
    if (!parsedData || !selectEl || !selectEl.options.length) {
        if(qtyDisplay) qtyDisplay.innerText = "0";
        return;
    }

    const colIdx = parseInt(selectEl.value);
    let total = 0;

    parsedData.body.forEach(row => {
        let rawVal = String(row[colIdx] || "");
        // Strip currency/text so only the number is summed.
        let num = parseFloat(rawVal.replace(/[^0-9.-]/g, ''));
        if (!isNaN(num)) total += num;
    });

    qtyDisplay.innerText = total.toLocaleString();
}

// Recompute the Step 1 live totals whenever either textarea changes.
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('tableA').addEventListener('input', () => scanForLiveQty('A'));
    document.getElementById('tableB').addEventListener('input', () => scanForLiveQty('B'));
});

// ==========================================
// STEP 2 LIVE QUANTITY (dropdown appears when there are 2+ qty columns)
// Re-runs on every Step 2 render, so editing a quantity cell updates the total live.
// ==========================================
function updateStep2Qty(side) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p?.dataA : p?.dataB;
    const qtyEl = document.getElementById(`step2Qty${side}`);
    const selectEl = document.getElementById(`step2QtySelect${side}`);

    if (!data || !data.headers) {
        if (qtyEl) qtyEl.innerText = "0";
        if (selectEl) selectEl.style.display = 'none';
        return;
    }

    // Find every quantity-like column (skip MATRIX rule columns)
    const qtyRegex = /qty|quantity|units|pcs|order/i;
    const qtyCols = [];
    data.headers.forEach((h, i) => {
        if (h && qtyRegex.test(String(h)) && !String(h).includes('[MATRIX]')) {
            qtyCols.push({ name: String(h), index: i });
        }
    });

    if (qtyCols.length === 0) {
        if (selectEl) selectEl.style.display = 'none';
        if (qtyEl) qtyEl.innerText = "0";
        return;
    }

    // Decide which column to total: remembered choice -> "adjusted qty" -> first
    if (!p.uiState) p.uiState = {};
    const storedIdx = p.uiState[`qtyCol${side}`];
    let chosen = qtyCols.find(c => c.index === storedIdx);
    if (!chosen) {
        chosen = qtyCols.find(c => /adjusted\s*qty|adjusted\s*quantity/i.test(c.name)) || qtyCols[0];
        p.uiState[`qtyCol${side}`] = chosen.index;
    }

    // Build the dropdown (only visible when the vendor gave us multiple qty columns)
    if (selectEl) {
        selectEl.innerHTML = "";
        qtyCols.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.index;
            opt.text = c.name.length > 22 ? c.name.slice(0, 20) + '…' : c.name;
            opt.title = c.name;
            selectEl.appendChild(opt);
        });
        selectEl.value = chosen.index;
        selectEl.style.display = qtyCols.length > 1 ? 'inline-block' : 'none';
    }

    // Sum the chosen column across all active rows
    let total = 0;
    data.body.forEach(row => {
        const num = parseFloat(String(row[chosen.index] || "").replace(/[^0-9.-]/g, ''));
        if (!isNaN(num)) total += num;
    });
    if (qtyEl) qtyEl.innerText = total.toLocaleString();
}

function onStep2QtyChange(side) {
    const p = projects[activeProjectIdx];
    if (!p) return;
    const selectEl = document.getElementById(`step2QtySelect${side}`);
    if (!p.uiState) p.uiState = {};
    p.uiState[`qtyCol${side}`] = parseInt(selectEl.value);
    updateStep2Qty(side);
}


// ==========================================
// DELETE / RESTORE A SET (with 30s undo)
// ==========================================

// Timer that hides the "Restore Set" button after the undo window elapses.
let restoreTimeoutId = null;

// Removes a Set, remembering it so it can be restored within 30 seconds.
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

        // After 30s, hide the button and forget the deleted Set for good.
        if (restoreTimeoutId) clearTimeout(restoreTimeoutId);
        restoreTimeoutId = setTimeout(() => {
            undoBtn.style.display = 'none';
            deletedItem = null;
        }, 30000);
    }

    if(isOverviewMode) showOverview();
    else switchProject(activeProjectIdx);
}

// Re-inserts the most recently deleted Set at its original position.
function restoreProject() {
    if (!deletedItem) return;

    // The user restored in time — cancel the auto-hide timer.
    if (restoreTimeoutId) clearTimeout(restoreTimeoutId);

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
// MODALS & ALERTS
// ==========================================

// Opens the shared modal. `type` ('success'/'error'/'confirm') picks the icon,
// colour and button label; `callback` runs on confirm.
function showModal(title, content, type = 'success', callback = null) {
    const modal = document.getElementById('customModal');
    const titleEl = document.getElementById('modalTitle');
    const msgEl = document.getElementById('modalMsg');
    const iconEl = document.getElementById('modalIcon');
    const btnEl = document.getElementById('modalBtn');

    if (callback) {
        btnEl.onclick = function () { callback(); closeModal(); };
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
        alert(title + "\n" + content.replace(/<br>/g, '\n'));
    }
}

function closeModal() {
    document.getElementById('customModal')?.classList.remove('open');
}

window.onclick = function (event) {
    if (event.target === document.getElementById('customModal')) closeModal();
}

// ==========================================
// TOP BAR (SET TABS) & OVERVIEW
// ==========================================

// Renders the tab strip, grouping the Sets that came from the same file into one
// "cuboid" box.
function renderTopBar() {
    const list = document.getElementById('projectTabs');
    list.innerHTML = "";

    let lastFileName = null;
    let html = "";

    projects.forEach((p, i) => {
        const fName = p.fileName || "";
        const activeClass = (i === activeProjectIdx && !isOverviewMode) ? 'active' : '';

        if (fName !== lastFileName) {
            // Close the previous file's box before starting a new one.
            if (lastFileName !== null) html += `</div>`;

            // Shorten very long file names for display (e.g. 19b734...28a3f.xlsx).
            let displayFileName = fName;
            if (fName.length > 25) {
                displayFileName = fName.substring(0, 8) + "..." + fName.substring(fName.length - 8);
            }

            // Open a new file box; label it by filename, or "Manual Data" if none.
            html += `<div class="file-cuboid" style="display:flex; align-items:center; background:#f8fafc; border:1px solid #cbd5e1; border-radius:6px; padding:3px; margin-right:8px; box-shadow: inset 0 1px 2px rgba(0,0,0,0.05);">`;

            if (fName !== "") {
                html += `<div style="padding:0 10px; font-size:11px; font-weight:700; color:#334155; display:flex; align-items:center; gap:6px;" title="${fName}">
                            <i class="fas fa-folder-open" style="color:#2563eb;"></i> ${displayFileName}
                         </div>`;
            } else {
                html += `<div style="padding:0 10px; font-size:11px; font-weight:700; color:#334155; display:flex; align-items:center; gap:6px;">
                            <i class="fas fa-edit" style="color:#f59e0b;"></i> Manual Data
                         </div>`;
            }
            lastFileName = fName;
        }

        // One tab per Set inside the current file box.
        html += `
            <div class="tab-item ${activeClass}" onclick="switchProject(${i})" style="margin:0 2px;">
                <div class="tab-dot ${p.status}"></div>
                <span>${p.name}</span>
                <button class="btn-tab-close" onclick="deleteProject(event, ${i})">×</button>
            </div>`;
    });

    if (lastFileName !== null) html += `</div>`; // close the last file box
    
    list.innerHTML = html;
    updateAddButtonText();
}

// Shows the Overview screen: a table of every Set plus roll-up totals across them.
function showOverview() {
    saveCurrentViewToProject();
    isOverviewMode = true;

    document.querySelector('.sidebar').classList.add('overview-locked');
    document.getElementById('tabOverview').classList.add('active');
    document.querySelectorAll('.tab-item').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById('overviewSection').style.display = 'block';

    // Tally per-Set status and match/mismatch counts.
    let totalSets = projects.length;
    let processedSets = 0;
    let pendingSets = 0;

    let totalRows = 0;
    let totalMatches = 0;
    let totalMismatches = 0;
    let tbody = "";

    projects.forEach((p, idx) => {
        if (p.status === 'done') processedSets++;
        else pendingSets++;

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

    // Roll-up stat cards across all Sets.
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

// Switches the active Set: saves the current view, resets selection state, and
// loads the chosen Set into the editor.
function switchProject(idx, skipSave = false) {
    if (isOverviewMode) {
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

// Writes whatever is currently on screen (textareas, matrix, mapping, result options)
// back into the active Set, so nothing is lost when switching Sets or steps.
function saveCurrentViewToProject() {
    if (projects.length === 0 || isOverviewMode || activeProjectIdx < 0) return;

    const p = projects[activeProjectIdx];
    if (!p.settings) p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };

    if (document.getElementById('step1').style.display !== 'none') {
        p.rawA = document.getElementById('tableA').value;
        p.rawB = document.getElementById('tableB').value;
        p.rawMatrix = document.getElementById('matrixRawInput').value;
        p.matrix = getMatrixDataFromUI(); 
        const matSec = document.getElementById('matrixSection');
        p.showMatrix = (matSec && matSec.style.display !== 'none');

        p.settings.modeA = document.getElementById('headerModeA')?.value || '1row';
        p.settings.modeB = document.getElementById('headerModeB')?.value || '1row';
    }

    // On Step 3, capture the current mapping choices too.
    if (document.getElementById('step3').style.display !== 'none') {
        if (typeof saveMappingFromUI === 'function') saveMappingFromUI();
    }

    if (document.getElementById('step4').style.display !== 'none') {
        p.settings.trimResults = document.getElementById('chkTrimResults')?.checked || false;
        p.settings.showErrors = document.getElementById('chkShowErrors')?.checked || false;
    }
}


// Loads a Set into the editor: fills the textareas, restores settings/matrix,
// and jumps to whichever step that Set was last on.
function loadProjectIntoView(idx) {
    const p = projects[idx];
    if (!p) return;

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
        matrixInput.oninput = function () { handleMatrixInput(this, idx); };
        matrixInput.onpaste = function (e) { handleMatrixPaste(e, idx); };
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

    // Restore this Set's own parsing modes and result toggles to the UI.
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
        toggleMismatchView();
    }

    jumpToStep(p.step || 1); // also renders steps 2, 3 and 4
    // Refresh the Step 1 live totals for the newly loaded Set.
    if (p.step === 1) { scanForLiveQty('A'); scanForLiveQty('B'); }
}

// Blanks out the active Set back to an empty Step 1 state.
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

    p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };

    const badge = document.getElementById('table2FileName');
    if (badge) badge.style.display = 'none';

    loadProjectIntoView(activeProjectIdx);
    renderTopBar();
}

// Shows the given step (1–4), rendering its content and updating the sidebar nav.
function jumpToStep(step) {
    const p = projects[activeProjectIdx];
    p.step = step;

    if (step === 2) {
        // Always (re)build Step 2 on entry — fixes blank tables when returning from Step 3/4.
        // Only re-parse from raw text if the cleaned data was dropped, so Step 2 edits survive.
        if (!p.dataA && p.rawA) {
            p.dataA = parseExcelData(p.rawA, p.settings?.modeA || '1row');
            if (p.dataA) autoCleanData(p.dataA);
        }
        if (!p.dataB && p.rawB) {
            p.dataB = parseExcelData(p.rawB, p.settings?.modeB || '1row');
            if (p.dataB) autoCleanData(p.dataB);
        }
        renderPreviewTables();
    }
    if (step === 3) renderMappingTable();
    if (step === 4) renderDashboard();

    document.querySelectorAll('.page-section').forEach(el => el.style.display = 'none');
    document.getElementById(`step${step}`).style.display = 'block';

    document.querySelectorAll('.v-step').forEach(el => el.classList.remove('active'));
    document.getElementById(`navStep${step}`).classList.add('active');
}

// ==========================================
// CPQ UPLOAD (Sainpase / Table 2 side)
// ==========================================

// Loads CPQ Excel files into existing Sets (one file per Set, in order),
// extracting the SKU table from sheet 2 when present.
function handleCPQUpload(input) {
    if (!input.files || input.files.length === 0) return;

    saveCurrentViewToProject();

    // Oldest-first so files line up with Sets in a predictable order.
    const files = Array.from(input.files).sort((a, b) => a.lastModified - b.lastModified);

    let processedCount = 0;
    const startIdx = activeProjectIdx;

    // Reads CPQ files sequentially, filling Sets starting at the active one.
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

        projects[targetIdx].fileNameB = file.name;

        // If we're filling the Set on screen, update its file badge live.
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

        reader.onload = function (e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                let extracted = "";

                // Prefer sheet 2's SKU table (drop its first column), if present.
                if (workbook.SheetNames.length > 1) {
                    const sheet2 = workbook.Sheets[workbook.SheetNames[1]];
                    const rows = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
                    let start = -1;
                    for (let r = 0; r < rows.length; r++) {
                        if ((rows[r] || []).join(" ").toLowerCase().includes("sku information")) {
                            start = r + 1;
                            break;
                        }
                    }
                    if (start !== -1) {
                        extracted = rows.slice(start)
                            .map(r => r.slice(1).map(c => (c == null) ? "" : c).join("\t"))
                            .join("\n");
                    }
                }

                // Otherwise fall back to the whole of sheet 1.
                if (!extracted) {
                    const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
                    extracted = rows.map(r => r.map(c => (c == null) ? "" : c).join("\t")).join("\n");
                }

                projects[targetIdx].rawB = extracted;
                projects[targetIdx].dataB = null; // force a re-parse on next render
                processedCount++;

            } catch (err) {
                console.error("Error reading file for Set " + (targetIdx + 1), err);
            } finally {
                processNext(i + 1);
            }
        };
        reader.readAsArrayBuffer(file);
    };

    processNext(0); // start the sequential read
}

// ==========================================
// STEP 1 -> STEP 2 (parse & clean)
// ==========================================

// Validates the Step 1 textareas, parses both sides (reusing cached data when the
// text is unchanged), folds in matrix rules, and moves to Step 2.
function goToPreview() {
    try {
        const inputA = document.getElementById('tableA');
        const inputB = document.getElementById('tableB');

        if (!inputA.value.trim() || !inputB.value.trim()) {
            return showModal("Missing Data", "Please paste data into both tables.", 'error');
        }

        const p = projects[activeProjectIdx];
        const modeA = document.getElementById('headerModeA').value || "1row";
        const modeB = document.getElementById('headerModeB').value || "1row";

        // Only re-parse a side when its text actually changed, so Step 2 edits survive.
        let changedA = String(p.rawA || "").trim() !== inputA.value.trim();
        let changedB = String(p.rawB || "").trim() !== inputB.value.trim();

        p.rawA = inputA.value;
        p.rawB = inputB.value;

        if (changedA || !p.dataA) {
            p.dataA = parseExcelData(p.rawA, modeA);
            if (p.dataA) autoCleanData(p.dataA);
        }
        if (changedB || !p.dataB) {
            p.dataB = parseExcelData(p.rawB, modeB);
            if (p.dataB) autoCleanData(p.dataB);
        }

        if (!p.dataA || !p.dataB) return showModal("Error Parsing", "Check input format.", 'error');

        applyMatrixColumns(p);

        p.status = 'ready';
        p.step = 2;
        renderTopBar();
        jumpToStep(2); // renders the Step 2 preview tables
    } catch (err) {
        showModal("System Error", err.message, 'error');
    }
}

// Renders one Step 2 editable preview table (headers, rows, action buttons, toolbar)
// for the given side, and refreshes its live quantity total.
function renderSinglePreview(containerId, countId, data, side) {
    if (!data || !data.body) return;

    document.getElementById(countId).innerText = data.body.length;
    const container = document.getElementById(containerId);

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

    // Refresh the live total (and column picker) for this side.
    updateStep2Qty(side);

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
    
    const stickyStyle = "min-height: 45px; padding: 8px 15px; display:flex; flex-wrap: wrap; gap: 8px; justify-content:space-between; align-items:center; position: sticky; top: 0; left: 0; background: #fff; z-index: 30; border-bottom: 1px solid #e2e8f0; width: 100%; box-sizing: border-box; box-shadow: 0 2px 4px rgba(0,0,0,0.05);";

    const hasTranslationBackup = !!p[`backupHeaders${side}`];
    const undoDisplay = hasTranslationBackup ? 'inline-block' : 'none';

    let html = `
    <div style="${stickyStyle}">
        <div style="display:flex; flex-wrap:wrap; gap:8px; align-items:center;">
            <button onclick="toggleHiddenTable('hiddenBox-${side}', this)" style="background:none; border:none; color:${btnColor}; font-size:12px; cursor:pointer; font-weight:600; display:flex; align-items:center; gap:5px;"><i class="fas ${btnIcon}"></i> ${btnText}</button>
            <div style="width:1px; height:14px; background:#cbd5e1; margin: 0 2px;"></div>
            <button onclick="showSplitColumnsModal('${side}')"
        style="background:#eff6ff;
               color:#2563eb;
               border:1px solid #bfdbfe;
               padding:4px 8px;
               border-radius:4px;
               font-size:11px;
               font-weight:600;
               cursor:pointer;">
    <i class="fas fa-code-branch"></i> Split
</button>
<button
    onclick="showSplitByRowsModal('${side}')"
    style="
        background:#f8fafc;
        color:#475569;
        border:1px solid #cbd5e1;
        padding:4px 8px;
        border-radius:4px;
        font-size:11px;
        font-weight:600;
        cursor:pointer;
    "
>
    <i class="fas fa-cut"></i> Split by Rows
</button>
<button
    onclick="showSplitByUniquePatternModal('${side}'); event.stopPropagation()"
    style="
        background:#faf5ff;
        color:#7c3aed;
        border:1px solid #ddd6fe;
        padding:4px 8px;
        border-radius:4px;
        font-size:11px;
        font-weight:600;
        cursor:pointer;
    "
>
    <i class="fas fa-layer-group"></i> Unique Pattern
</button>
            <button onclick="stitchColumns('${side}')" style="background:#f8fafc; color:#475569; border:1px solid #cbd5e1; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer; transition:0.2s;" onmouseover="this.style.background='#e2e8f0'"><i class="fas fa-link"></i> Merge</button>
            <button onclick="bulkTextEdit('${side}', 'add')" style="background:#f8fafc; color:#10b981; border:1px solid #a7f3d0; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;"><i class="fas fa-plus"></i> Text</button>
            <button onclick="bulkTextEdit('${side}', 'remove')" style="background:#f8fafc; color:#ef4444; border:1px solid #fecaca; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;"><i class="fas fa-minus"></i> Text</button>
            <div style="width:1px; height:14px; background:#cbd5e1; margin: 0 2px;"></div>
            <button onclick="universalTranslate('${side}')" style="background:#fffbeb; color:#d97706; border:1px solid #fde68a; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;"><i class="fas fa-globe"></i> Translate</button>
            <button id="undoTranslateBtn${side}" onclick="undoTranslation('${side}')" style="display:${undoDisplay}; background:#fef2f2; color:#ef4444; border:1px solid #fca5a5; padding:4px 8px; border-radius:4px; font-size:11px; font-weight:600; cursor:pointer;"><i class="fas fa-undo"></i> Undo</button>
        </div>
        <div style="display:flex; align-items:center; gap:10px;">
            <span style="font-size:11px; color:#64748b; font-weight:600; display:none;" id="undoBadge${side}"><i class="fas fa-undo"></i> Ctrl+Z to Undo</span>
            <button onclick="addNewRow('${side}')" style="background:#10b981; color:white; border:none; padding:6px 12px; border-radius:4px; font-size:11px; font-weight:bold; cursor:pointer;"><i class="fas fa-plus"></i> Add Row</button>
        </div>
    </div>
    <div style="outline:none; position:relative;" tabindex="0" id="gridContainer-${side}">
        <table class="clean-table" id="previewTable${side}" style="user-select:none; outline:none; border-collapse: separate; border-spacing: 0;">
            <thead><tr><th style="width:50px;">Action</th><th style="width:40px;">#</th>`;

    data.headers.forEach((h, i) => {
        let isMatrix = String(h).includes('[MATRIX]');
        let headerClass = isMatrix ? "clickable-head matrix-header" : "clickable-head";
        html += `<th class="${headerClass}" data-c="${i}">${h}</th>`;
    });

    html += "</tr></thead><tbody>";

    displayList.forEach((item) => {
        const bgHex = item.type === 'hidden' ? "#f1f5f9" : "#ffffff";
        const txtHex = item.type === 'hidden' ? "#94a3b8" : "inherit";
        const actionBtn = item.type === 'active'
            ? `<button class="btn-ghost" style="color:#ef4444;" onclick="deleteRow('${side}', ${item.realIndex})"><i class="fas fa-trash"></i></button>`
            : `<button class="btn-ghost" style="color:#10b981;" onclick="restoreHiddenRow('${side}', ${item.realIndex})"><i class="fas fa-plus-circle"></i></button>`;

        html += `<tr style="background-color: ${bgHex}; color: ${txtHex};" data-r="${item.realIndex}">
            <td style="text-align:center;">${actionBtn}</td>
            <td class="row-num" style="font-size:11px; color:#aaa; text-align:center; cursor:pointer;">${item.originalId + 1}</td>`;

        item.row.forEach((cell, cIdx) => {
            let isMatrix = String(data.headers[cIdx]).includes('[MATRIX]');
            let cellClass = isMatrix ? "matrix-val" : (side === 'B' ? "table2-val" : "");
            html += `<td data-r="${item.realIndex}" data-c="${cIdx}" class="${cellClass}">${cell}</td>`;
        });
        html += `</tr>`;
    });

    html += `</tbody></table><div id="fillHandle${side}" class="fill-handle" style="display:none;"></div></div>`;
    container.innerHTML = html;
    setTimeout(() => { GridEngine.init(side); }, 50);
}


// Renders both Step 2 preview tables (Form and CPQ), clearing a side that has no data.
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
    if (deletedRowsHistory.length > 0) {
        undoBtn.style.display = 'inline-flex';
        undoBtn.innerText = `Undo Delete (${deletedRowsHistory.length})`;
    } else {
        undoBtn.style.display = 'none';
    }
}

// Current cell/column selection state for the Step 2 spreadsheet grid.
let excelState = {
    side: null,      // 'A' or 'B'
    mode: null,      // 'cell' or 'col'
    r: -1,           // active row index
    c: -1,           // active column index
    editing: false   // true while the user is typing in a cell
};



// Moves a Step 2 row into the "hidden/ignored" list (reversible), then re-renders.
function deleteRow(side, rowIndex) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;

    const rowToDelete = dataObj.body[rowIndex];
    if (!dataObj.hiddenRows) dataObj.hiddenRows = [];

    // Remember the row's original id so it can be restored to its place later.
    const originalId = (rowToDelete._originalIdx !== undefined) ? rowToDelete._originalIdx : rowIndex;

    dataObj.hiddenRows.unshift({ data: rowToDelete, restoreIdx: originalId });
    dataObj.body.splice(rowIndex, 1);

    renderPreviewTables();
}

// Hides a range of rows (1-based "start-end") from the range input box.
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
    // Iterate high-to-low so splicing doesn't shift the indices we still need.
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
// Puts back the most recently row deleted via the keyboard-delete history.
function restoreLastRow() {
    if (deletedRowsHistory.length === 0) return;

    const last = deletedRowsHistory.pop();
    const p = projects[activeProjectIdx];

    if (last.side === 'A') p.dataA.body.splice(last.idx, 0, last.data);
    else p.dataB.body.splice(last.idx, 0, last.data);

    renderPreviewTables();
}

// ==========================================
// STEP 3: COLUMN MAPPING
// ==========================================

// Guards against unequal row counts, then opens Step 3.
function goToMapping() {
    const p = projects[activeProjectIdx];
    if (p.dataA.body.length !== p.dataB.body.length) {
        showModal("Row Mismatch", `Form Data: ${p.dataA.body.length} rows<br>Sainpase Data: ${p.dataB.body.length} rows<br><br>Row counts must be equal.`, 'error');
        return;
    }
    renderMappingTable();
    jumpToStep(3);
}

// Builds the Step 3 mapping grid: one row per Form/Matrix column with a dropdown of
// CPQ columns, auto-suggesting matches (exact name, synonym group, or saved memory).
function renderMappingTable() {
    const p = projects[activeProjectIdx];
    if (!p.dataA || !p.dataB) return false;

    const tbody = document.getElementById('mappingBody');
    tbody.innerHTML = "";
    document.getElementById('mapSearchInput').value = "";

    const leftRows = [];
    // Matrix rule names, used to spot Form columns that duplicate a matrix rule.
    const matrixKeys = p.dataA.headers
        .filter(h => h && h.startsWith('[MATRIX] '))
        .map(h => h.replace('[MATRIX] ', '').toLowerCase().trim());

    p.dataA.headers.forEach((h, i) => {
        if (h && h.trim() !== "") {
            let isMatrix = h.startsWith('[MATRIX] ');
            let displayLabel = isMatrix ? h.replace('[MATRIX] ', '') : h;

            // Matrix / table name clash: keep the table column visible but flag it so it
            // defaults to unmapped — the matrix rule is used unless the user deliberately
            // maps this table column and unmaps the matrix one.
            let dupOfMatrix = !isMatrix && matrixKeys.includes(displayLabel.toLowerCase().trim());

            // "Total Qty" is a summary column, never mapped — omit it from the grid.
            if (!isMatrix && (displayLabel.toLowerCase() === 'total qty' || displayLabel.toLowerCase() === 'total quantity')) return;

            leftRows.push({ type: 'source', name: h, display: displayLabel, val: i, dupOfMatrix: dupOfMatrix });
        }
    });

    const rightOptions = p.dataB.headers
        .map((h, i) => ({ name: h, index: i }))
        .filter(opt => {
            if (!opt.name || opt.name.trim() === "") return false;
            const cleanName = opt.name.toLowerCase().trim();
            return !cleanName.startsWith("unnamed") && cleanName !== "adjusted quantity" && cleanName !== "layout1";
        });
    rightOptions.sort((a,b) => a.name.localeCompare(b.name));

    // Maps a header to a synonym group (e.g. "qty"/"quantity" -> QTY_GROUP) so
    // differently-named-but-equivalent columns still auto-match.
    function getStandardKey(headerName) {
        let clean = headerName.toLowerCase().replace(/\(?\boptional\b\)?/g, "").replace(/[\u4e00-\u9fff]/g, "").replace(/[^a-z0-9]/g, "").trim();
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

        // Prefer a mapping the user already saved for this Set.
        const savedSession = p.mapping.find(m => {
            return (row.type === 'source' && m.targetType === 'source' && m.targetVal === row.val) ||
                   (row.type === 'matrix' && m.targetType === 'matrix' && m.targetVal === row.val);
        });

        if (savedSession) {
            matchVal = savedSession.idxB;
        } else if (!p.isMapped && !row.dupOfMatrix) {
            // Leave table columns that clash with a matrix rule unmapped by default
            const cleanSource = row.name.toLowerCase().trim();
            const cleanDisplay = row.display ? row.display.toLowerCase().trim() : cleanSource;
            
            let directMatch = rightOptions.find(opt => opt.name.toLowerCase().trim() === cleanSource);
            if (!directMatch) directMatch = rightOptions.find(opt => opt.name.toLowerCase().trim() === cleanSource || opt.name.toLowerCase().trim() === cleanDisplay);

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
                        return (row.type === 'source' && sType === 'SRC' && sVal === row.name) || (row.type === 'matrix' && sType === 'MAT' && sVal === row.name);
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
        
        let label = row.display;
        if (row.name.startsWith('[MATRIX] ')) {
            let mRule = p.matrix.find(m => m.key.replace(/^\*/, '').toLowerCase() === row.display.toLowerCase());
            let mVal = mRule ? mRule.val : "";
            label = `${row.display} <small style="color:#666">(${mVal})</small>`;
        } else if (row.dupOfMatrix) {
            // Explain why this table column starts unmapped
            label = `${row.display} <small style="color:#f59e0b">(table copy — Matrix rule in use)</small>`;
        }
        
        const tr = document.createElement('tr');
        tr.className = rowClass;
        tr.setAttribute('data-search', row.name.toLowerCase());
        
        tr.innerHTML = `
            <td><b>${label}</b></td>
            <td style="text-align:center"><i class="fas fa-arrow-right" style="color:#9ca3af"></i></td>
            <td style="display:flex; align-items:center; gap:8px;">
                <select style="flex:1;" class="map-select" data-type="${row.type}" data-val="${row.val}" data-name="${row.name}" onchange="autoTick(this)">${opts}</select>
                <button onclick="forgetSingleRule(this)" style="background:none; border:none; color:#ef4444; cursor:pointer; font-size:14px; opacity:0.6;"><i class="fas fa-eraser"></i></button>
            </td>
            <td style="text-align:center"><input type="checkbox" class="map-check" ${isChecked} onchange="updateMapStats()"></td>`;
        
        tbody.appendChild(tr);
    });
    
    updateMapOptionsVisibility();
    updateMapStats();
    return true;
}

// Hides CPQ options already chosen by another row, enforcing one target per source.
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

// When a mapping dropdown changes: tick/untick its "include" box and remember the
// choice in localStorage so the same columns auto-map next time.
function autoTick(selectEl) {
    const checkbox = selectEl.parentElement.nextElementSibling.querySelector('input');

    if (selectEl.value !== "-1") {
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

// Updates the "All mapped" / "Unmapped: ..." banner from the current checkbox states.
function updateMapStats() {
    const rows = document.querySelectorAll('#mappingBody tr');
    let mappedCount = 0;
    let unmappedNames = [];

    rows.forEach(r => {
        const checkbox = r.querySelector('.map-check');
        const nameLabel = r.querySelector('td:first-child b');

        if (checkbox && checkbox.checked) {
            mappedCount++;
        } else if (nameLabel) {
            let cleanName = nameLabel.innerText.replace(/\s*\(.*?\)/, '').trim();
            unmappedNames.push(cleanName);
        }
    });

    const unmappedCount = rows.length - mappedCount;
    const alertEl = document.getElementById('unmappedList');

    if (alertEl) {
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

// Clears every mapping dropdown/checkbox in the current grid (UI only).
function resetMapping() {
    document.querySelectorAll('.map-select').forEach(s => { s.value = "-1"; });
    document.querySelectorAll('#mappingBody tr').forEach(r => {
        r.classList.remove('mapped-row');
        r.querySelector('.map-check').checked = false;
    });
    updateMapOptionsVisibility();
    updateMapStats();
}

function resetViewOnly() {
    resetMapping();
    showToast("View Reset");
}

// Forgets any saved mapping for this Set and re-runs the auto-mapper from scratch.
function autoMapAgain() {
    const p = projects[activeProjectIdx];
    if (!p) return;

    p.isMapped = false;
    p.mapping = [];

    resetMapping();
    renderMappingTable(); // re-runs the auto-suggestion logic

    showToast("Auto-Map Applied");
}

// Filters the mapping rows by the search box (matches the source column name).
function filterMappingRows() {
    const filter = document.getElementById('mapSearchInput').value.toLowerCase();
    document.querySelectorAll('#mappingBody tr').forEach(r => {
        r.style.display = (r.getAttribute('data-search') || "").includes(filter) ? "" : "none";
    });
}

// Forgets all learned mapping rules for this Set's CPQ columns.
function clearSavedRules() {
    showModal("Clear Saved Rules?", "This will forget all auto-learned mappings.", "confirm", () => {
        const p = projects[activeProjectIdx];
        if (p.dataB) p.dataB.headers.forEach(h => localStorage.removeItem("map_" + h));
        resetMapping();
        showToast("Memory Cleared.");
    });
}

// Forgets the learned rule for one mapping row (the eraser button) and clears it.
function forgetSingleRule(btnElement) {
    const tr = btnElement.closest('tr');
    const selectEl = tr.querySelector('.map-select');

    if (selectEl.value !== "-1") {
        const targetName = selectEl.options[selectEl.selectedIndex].text;

        localStorage.removeItem("map_" + targetName);

        // Clear this row and refresh stats/visibility.
        selectEl.value = "-1";
        autoTick(selectEl);

        showToast(`Forgot memory for "${targetName}"`);
    } else {
        showToast(`No rule to forget here.`);
    }
}

// Clones the current mapping to every other Set BY COLUMN NAME.
// (Cloning by raw column index is wrong when other Sets have a different column
// order — it would silently map the wrong fields. Name-matching fixes that.)
function applyMappingToAll() {
    // 1. Force save the current screen's mapping first
    saveMappingFromUI();
    const currentP = projects[activeProjectIdx];

    if (!currentP.mapping || currentP.mapping.length === 0) {
        return showModal("No Mapping", "Please map at least one column first before copying.", "warning");
    }

    // Case-insensitive header lookup
    const findHeader = (headers, name) => {
        if (!headers || name == null) return -1;
        const t = String(name).toLowerCase().trim();
        return headers.findIndex(h => String(h).toLowerCase().trim() === t);
    };

    // 2. Turn the current mapping into NAME pairs (source header name -> target header name)
    const namePairs = currentP.mapping.map(m => {
        const sourceName = (m.targetType === 'source')
            ? currentP.dataA?.headers?.[m.targetVal]
            : m.targetVal; // matrix key
        return { sourceName, targetName: m.name, targetType: m.targetType };
    }).filter(pair => pair.sourceName != null && pair.targetName != null);

    let appliedCount = 0;
    const partial = [];  // sets where some columns couldn't be matched
    const skipped = [];  // sets that shared no columns at all (left untouched)

    projects.forEach((p, idx) => {
        if (idx === activeProjectIdx || p.status === 'empty') return;

        // Make sure this set is parsed so we can match by name
        if (!p.dataA && p.rawA) {
            p.dataA = parseExcelData(p.rawA, p.settings?.modeA || '1row');
            if (p.dataA) autoCleanData(p.dataA);
        }
        if (!p.dataB && p.rawB) {
            p.dataB = parseExcelData(p.rawB, p.settings?.modeB || '1row');
            if (p.dataB) autoCleanData(p.dataB);
        }
        if (!p.dataA || !p.dataB) { skipped.push(p.name); return; }

        // Fold in this set's Matrix columns so matrix mappings resolve (matches manual)
        applyMatrixColumns(p);

        const newMapping = [];
        const usedTargets = new Set();
        let missed = 0;

        namePairs.forEach(pair => {
            let srcIdx = findHeader(p.dataA.headers, pair.sourceName);
            // Matrix columns may or may not carry the "[MATRIX] " prefix in each set
            if (srcIdx === -1 && pair.targetType === 'matrix') {
                srcIdx = findHeader(p.dataA.headers, '[MATRIX] ' + pair.sourceName);
            }
            const tgtIdx = findHeader(p.dataB.headers, pair.targetName);

            if (srcIdx === -1 || tgtIdx === -1 || usedTargets.has(tgtIdx)) { missed++; return; }

            usedTargets.add(tgtIdx);
            newMapping.push({
                name: p.dataB.headers[tgtIdx],
                idxB: tgtIdx,
                targetType: 'source',
                targetVal: srcIdx
            });
        });

        // Also exact-name map any EXTRA source columns this Set has that the base
        // Set's mapping didn't cover. If table 2 has a column with the exact same
        // name, map it; if not, leave that column blank (unmapped).
        const matrixKeysP = p.dataA.headers
            .filter(h => String(h).startsWith('[MATRIX] '))
            .map(h => h.replace(/\[MATRIX\]\s*/i, '').toLowerCase().trim());

        p.dataA.headers.forEach((h, i) => {
            if (newMapping.some(m => m.targetVal === i)) return; // already mapped
            const isMatrixCol = String(h).startsWith('[MATRIX] ');
            const cleanName = String(h).replace(/\[MATRIX\]\s*/i, '').trim();
            // Matrix priority: skip a plain column that duplicates a matrix rule name
            if (!isMatrixCol && matrixKeysP.includes(cleanName.toLowerCase())) return;

            const tgtIdx = findHeader(p.dataB.headers, cleanName);
            if (tgtIdx === -1 || usedTargets.has(tgtIdx)) return; // no exact match -> leave blank

            usedTargets.add(tgtIdx);
            newMapping.push({
                name: p.dataB.headers[tgtIdx],
                idxB: tgtIdx,
                targetType: 'source',
                targetVal: i
            });
        });

        // Don't wipe a set's mapping if nothing matched — leave it as-is
        if (newMapping.length === 0) { skipped.push(p.name); return; }

        p.mapping = newMapping;
        p.isMapped = true;
        appliedCount++;
        if (missed > 0) partial.push(`${p.name} — ${missed} column(s) not found`);
    });

    // If we're re-viewing one of the sets we just changed, refresh the grid
    if (projects[activeProjectIdx] && projects[activeProjectIdx].step === 3) renderMappingTable();

    let msg = `✅ Applied your mapping to <strong>${appliedCount}</strong> other set(s), matched by column name.`;
    if (partial.length > 0) {
        msg += `<br><br><span style="color:#b45309;"><b>Partially applied</b> (some columns were missing):</span><br>${partial.join('<br>')}`;
    }
    if (skipped.length > 0) {
        msg += `<br><br><span style="color:#64748b;"><b>Skipped</b> (no matching columns — left unchanged):</span><br>${skipped.join('<br>')}`;
    }
    showModal("Mapping Cloned", msg, appliedCount > 0 ? "success" : "warning");
}


// ==========================================
// STEP 4: ANALYSIS & DASHBOARD
// ==========================================

// Validates the mapping (rows equal, at least one mapped, no duplicate targets),
// then computes and shows the comparison results for the active Set.
function generateResults() {
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

        // Reject if two source columns map to the same CPQ column.
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

// Reads the Step 3 grid into p.mapping, keeping only rows that are checked and
// point at a real CPQ column. Marks the Set as user-mapped.
function saveMappingFromUI() {
    const p = projects[activeProjectIdx];
    p.mapping = [];
    p.isMapped = true;

    // Read each row (not two parallel lists) so source/target indices can't drift.
    const rows = document.querySelectorAll('#mappingBody tr');

    rows.forEach(row => {
        const sel = row.querySelector('.map-select');
        const check = row.querySelector('.map-check');

        if (check && check.checked && sel && sel.value !== "-1") {
            const targetIdx = parseInt(sel.value);
            const srcType = sel.getAttribute('data-type');
            const srcVal = sel.getAttribute('data-val');

            // Only record the mapping if the target column actually exists.
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

// "Clean Data" (sidebar): parses and auto-cleans every Set, moving each to Step 2.
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

// "Run Analysis" (sidebar): parses, auto-maps and scores every Set in one pass,
// using the same auto-mapping as the manual Step 3 screen.
function runAllComparisons() {
    saveCurrentViewToProject();

    if (document.getElementById('step3').style.display !== 'none') {
        saveMappingFromUI();
    }

    let processed = 0;
    let mismatchCount = 0;
    let missingDataLog = [];

    projects.forEach((p, idx) => {
        // Use each Set's own parsing/trim settings.
        if (!p.settings) p.settings = { modeA: '1row', modeB: '1row', trimResults: false, showErrors: false };
        const modeA = p.settings.modeA;
        const modeB = p.settings.modeB;
        const doTrim = p.settings.trimResults;

        if (p.rawA && !p.dataA) {
            p.dataA = parseExcelData(p.rawA, modeA);
            if (p.dataA) autoCleanData(p.dataA);
        }
        if (p.rawB && !p.dataB) {
            p.dataB = parseExcelData(p.rawB, modeB);
            if (p.dataB) autoCleanData(p.dataB);
        }

        // Match manual: fold in the Matrix rule columns before mapping
        applyMatrixColumns(p);

        const hasA = !!(p.dataA && p.dataA.body.length > 0);
        const hasB = !!(p.dataB && p.dataB.body.length > 0);

        if (hasA && hasB) {
            if (p.dataA.body.length === p.dataB.body.length) {

                if (!p.mapping || p.mapping.length === 0) {
                    // Use the SAME auto-mapping as the manual Step 3 screen
                    autoMapLikeManual(idx);
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

// Appends the project's Matrix rules as "[MATRIX] key" columns onto dataA
// (same logic the manual Step 2 flow uses, extracted so bulk Run Analysis matches it).
function applyMatrixColumns(p) {
    if (!p || !p.matrix || p.matrix.length === 0 || !p.dataA || !p.dataA.headers) return;
    p.matrix.forEach(rule => {
        let key = rule.key.replace(/^\*/, ''); // Strip universal asterisk
        let val = rule.val.trim();
        if (val.toUpperCase() === "NO") val = "FALSE";
        if (val.toUpperCase() === "YES") val = "TRUE";

        let headerName = `[MATRIX] ${key}`;
        if (!p.dataA.headers.includes(headerName)) {
            p.dataA.headers.push(headerName);
            p.dataA.body.forEach(row => row.push(val));
            if (p.dataA.hiddenRows) p.dataA.hiddenRows.forEach(hRow => hRow.data.push(val));
        }
    });
}

// Builds a set's mapping using the EXACT same code path as the manual Step 3 screen,
// so bulk "Run Analysis" produces identical mappings to mapping each set by hand.
function autoMapLikeManual(idx) {
    const prev = activeProjectIdx;
    activeProjectIdx = idx;
    try {
        renderMappingTable();  // auto-suggests selects/checks exactly like manual Step 3
        saveMappingFromUI();   // commits those suggestions into projects[idx].mapping
    } finally {
        activeProjectIdx = prev;
    }
}

// Legacy standalone auto-mapper (no longer called — bulk Run Analysis now uses
// autoMapLikeManual so it matches the manual screen). Kept for reference.
function autoMapProject(p) {
    p.mapping = [];
    const usedTargets = new Set();

    // Maps a header to a broad synonym group so equivalent columns match.
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

    // Matrix rules win over a same-named plain column, so collect the matrix names.
    const matrixKeys = p.dataA.headers
        .filter(h => h.startsWith('[MATRIX] '))
        .map(h => h.replace(/\[MATRIX\]\s*/i, '').toLowerCase().trim());

    // Pass 1: exact name / remembered-rule matches.
    p.dataA.headers.forEach((h, i) => {
        let cleanHeader = h.replace(/\[MATRIX\]\s*/i, '').trim();
        let isMatrixCol = h.startsWith('[MATRIX] ');

        // Skip a plain column that duplicates a matrix rule name.
        if (!isMatrixCol && matrixKeys.includes(cleanHeader.toLowerCase())) {
            return;
        }

        let exactMatch = rightOptions.find(opt => !usedTargets.has(opt.index) && opt.name.toLowerCase().trim() === cleanHeader.toLowerCase());

        if (!exactMatch) {
            const memMatch = rightOptions.find(opt => {
                const saved = localStorage.getItem("map_" + opt.name);
                if (!saved) return false;
                const [sType, sVal] = saved.split(':');
                return sType === 'SRC' && (sVal === h || sVal === cleanHeader);
            });
            if (memMatch) exactMatch = memMatch;
        }

        if (exactMatch && !usedTargets.has(exactMatch.index)) {
            p.mapping.push({ name: exactMatch.name, idxB: exactMatch.index, targetType: 'source', targetVal: i });
            usedTargets.add(exactMatch.index);
        }
    });

    // Pass 2: fuzzy/synonym matches for whatever is still unmapped.
    p.dataA.headers.forEach((h, i) => {
        let cleanHeader = h.replace(/\[MATRIX\]\s*/i, '').trim();
        let isMatrixCol = h.startsWith('[MATRIX] ');

        if (!isMatrixCol && matrixKeys.includes(cleanHeader.toLowerCase())) {
            return;
        }

        const isMapped = p.mapping.some(m => m.targetType === 'source' && m.targetVal === i);
        if (!isMapped) {
            tryMap(cleanHeader, i, 'source');
        }
    });
}


// Builds the Step 4 summary modal: a perfect-match message, or a per-column table
// of mismatch counts.
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

// "Show Mismatches Only" toggle: hides perfect rows via an injected style rule
// and remembers the preference on the Set.
function toggleMismatchView() {
    const chk = document.getElementById('chkShowErrors');
    if (!chk) return;

    const onlyErrors = chk.checked;
    const styleId = 'mismatch-style';
    let styleTag = document.getElementById(styleId);

    if (onlyErrors) {
        if (!styleTag) {
            styleTag = document.createElement('style');
            styleTag.id = styleId;
            document.head.appendChild(styleTag);
        }
        styleTag.innerHTML = `.perfect-row { display: none !important; }`;
    } else {
        if (styleTag) styleTag.remove();
    }

    const p = projects[activeProjectIdx];
    if (p && p.settings) {
        p.settings.showErrors = onlyErrors;
    }
}

// Normalizes a value for comparison. Prices round to 2 decimals, quantities strip
// formatting to a plain number, and text is lower-cased while keeping exact spacing.
function sanitizeForComparison(val, type) {
    if (val === null || val === undefined) val = "";
    let str = String(val).replace(/[\r\n\t]/g, " ").trim();

    if (type === 'price') {
        let num = parseFloat(str.replace(/[^0-9.-]/g, ''));
        return isNaN(num) ? str.toLowerCase() : num.toFixed(2);
    } else if (type === 'qty') {
        let num = parseFloat(str.replace(/[^0-9.-]/g, ''));
        return isNaN(num) ? str.toLowerCase() : num.toString();
    } else {
        return str.toLowerCase();
    }
}

// Compares two values by field type, treating YES/TRUE and NO/FALSE as equal.
function checkEqualAdvanced(rawA, rawB, fieldName, doTrim) {
    const lowerName = fieldName.toLowerCase();
    const isPrice = /price|cost|retail/i.test(lowerName);
    const isQty = /qty|quantity|units|pcs|round/i.test(lowerName);
    
    let type = 'text';
    if (isPrice) type = 'price';
    else if (isQty) type = 'qty';

    let normA = sanitizeForComparison(rawA, type);
    let normB = sanitizeForComparison(rawB, type);
    
    if (doTrim && type === 'text') {
        normA = normA.replace(/\s+/g, '');
        normB = normB.replace(/\s+/g, '');
    }

    if (normA === normB) return { match: true, type: type };

    // Treat YES/TRUE (and NO/FALSE) as the same value.
    if (type === 'text') {
        const trueVals = ['yes', 'true'];
        const falseVals = ['no', 'false'];
        
        if (trueVals.includes(normA) && trueVals.includes(normB)) return { match: true, type: type };
        if (falseVals.includes(normA) && falseVals.includes(normB)) return { match: true, type: type };
    }

    return { match: false, type: type };
}

// Counts total matches/mismatches across every mapped column. A price with a value
// but no currency is treated as a mismatch.
function calculateStats(p, doTrim) {
    if (!p.dataA || !p.dataB || !p.mapping) return { matches: 0, mismatches: 0 };
    let matches = 0, mismatches = 0;
    const rows = Math.max(p.dataA.body.length, p.dataB.body.length);

    const currIdxB = p.dataB.headers.findIndex(h => /currency/i.test(h));

    p.mapping.forEach(map => {
        const isPrice = /price|cost|retail/i.test(map.name.toLowerCase());
        for(let i=0; i<rows; i++) {
            let vB = (p.dataB.body[i]?.[map.idxB] || "").toString();
            let vA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").toString();

            let result = checkEqualAdvanced(vA, vB, map.name, doTrim);

            if (isPrice) {
                let numB = parseFloat(vB.replace(/[^0-9.-]/g, ''));
                let currB = currIdxB !== -1 ? String(p.dataB.body[i]?.[currIdxB]).trim() : "";
                if (!isNaN(numB) && numB > 0 && currB === "") result.match = false;
            }

            if(result.match) matches++; else mismatches++;
        }
    });
    return { matches, mismatches };
}

// Renders the Step 4 dashboard: summary stats, per-column mismatch cards, and the
// side-by-side result tables. Shows a blank state when a side has no data.
function renderDashboard() { 
    const p = projects[activeProjectIdx]; 
    const doTrim = document.getElementById('chkTrimResults')?.checked || false;
    
    if (!p.dataA || !p.dataB) {
        const diffCards = document.getElementById('diffCards');
        if (diffCards) diffCards.innerHTML = `<div class="field-card" style="border-left: 4px solid #ccc; width:100%"><div class="fc-head">Demo Mode</div><div class="fc-stats" style="color:#666">No Data Loaded Yet</div></div>`;
        document.getElementById('globalStats').innerHTML = `<div class="big-stat"><div class="bs-val" style="color:#ccc">0</div><div class="bs-lbl">Rows</div></div>`;
        // Clear the column-health bar so it doesn't show another set's stale numbers
        const banner = document.getElementById('colHealthBanner');
        if (banner) banner.innerHTML = "";
        renderResultTables(0, doTrim);
        return;
    }

    const stats = calculateStats(p, doTrim);
    p.summary = stats;
    const maxRows = Math.max(p.dataA.body.length, p.dataB.body.length);

    // Column-health tallies: how many populated columns exist and how many are mapped.
    const colHasData = (dataObj, colIdx) => {
        if (!dataObj || !dataObj.body) return false;
        return dataObj.body.some(row => row[colIdx] !== null && row[colIdx] !== undefined && String(row[colIdx]).trim() !== "");
    };

    let activeFormCols = 0, mappedFormCols = 0;
    p.dataA.headers.forEach((h, i) => {
        if (colHasData(p.dataA, i)) {
            activeFormCols++;
            if (p.mapping.some(m => m.targetType === 'source' && m.targetVal === i)) mappedFormCols++;
        }
    });
    const unmappedFormCols = activeFormCols - mappedFormCols;

    const ignoreCPQList = ['adjusted quantity', 'layout1'];
    let activeCPQCols = 0, mappedCPQCols = 0;
    p.dataB.headers.forEach((h, i) => {
        const cleanHeader = String(h || "").toLowerCase().trim();
        if (!ignoreCPQList.includes(cleanHeader) && colHasData(p.dataB, i)) {
            activeCPQCols++;
            if (p.mapping.some(m => m.idxB === i)) mappedCPQCols++;
        }
    });
    const unmappedCPQCols = activeCPQCols - mappedCPQCols;
    const totalMapped = p.mapping.length;

    // Build the per-column mismatch cards.
    let diffCards = document.getElementById('diffCards');
    if (!diffCards) {
        diffCards = document.createElement('div');
        diffCards.id = 'diffCards';
        diffCards.className = 'cards-grid';
        const tablesGrid = document.querySelector('.tables-grid');
        if (tablesGrid && tablesGrid.parentNode) tablesGrid.parentNode.insertBefore(diffCards, tablesGrid);
    }
    diffCards.innerHTML = ""; 
    let hasMismatches = false;
    
    const currIdxB = p.dataB.headers.findIndex(h => /currency/i.test(h));

    p.mapping.forEach(map => { 
        let match = 0, miss = 0; 
        const isPrice = /price|cost|retail/i.test(map.name.toLowerCase());
        for (let i = 0; i < maxRows; i++) { 
            let vB = (p.dataB.body[i]?.[map.idxB] || "").toString(); 
            let vA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").toString(); 
            let result = checkEqualAdvanced(vA, vB, map.name, doTrim);
            let equal = result.match;
            if (isPrice) {
                let numB = parseFloat(vB.replace(/[^0-9.-]/g, ''));
                let currB = currIdxB !== -1 ? String(p.dataB.body[i]?.[currIdxB]).trim() : "";
                if (!isNaN(numB) && numB > 0 && currB === "") equal = false; 
            }
            if (equal) match++; else miss++; 
        } 
        if (miss > 0) {
            hasMismatches = true;
            diffCards.innerHTML += `<div class="field-card bg-warn"><div class="fc-head">${map.name}</div><div class="fc-stats"><span style="color:#10b981">✓ ${match}</span><span style="color:#ef4444">✗ ${miss}</span></div></div>`; 
        }
    }); 
    
    if (!hasMismatches) {
        diffCards.innerHTML = `<div style="grid-column: 1 / -1; text-align: center; padding: 40px; background: white; border-radius: 12px; border: 1px solid #bbf7d0; display:flex; flex-direction:column; align-items:center;"><i class="fas fa-check-circle" style="font-size: 48px; color: #22c55e; margin-bottom: 15px;"></i><h3 style="margin: 0; color: #15803d; font-size: 20px;">All Fields Matched!</h3><p style="margin: 10px 0 0 0; color: #64748b;">No mismatches found in the mapped columns.</p></div>`;
    }
    
    // Sum the mapped quantity column on both sides and warn loudly if they differ.
    let sumQtyA = 0;
    let sumQtyB = 0;
    let hasQtyCol = false;

    const qtyMap = p.mapping.find(m => /qty|quantity|units|pcs|round/i.test(m.name));

    if (qtyMap) {
        hasQtyCol = true;
        for (let i = 0; i < maxRows; i++) {
            let rawB = (p.dataB.body[i]?.[qtyMap.idxB] || "").toString();
            let rawA = qtyMap.targetType === 'matrix' ?
                (p.matrix.find(m => m.key === qtyMap.targetVal)?.val || "") :
                (p.dataA.body[i]?.[qtyMap.targetVal] || "").toString();

            let numA = parseFloat(rawA.replace(/[^0-9.-]/g, ''));
            let numB = parseFloat(rawB.replace(/[^0-9.-]/g, ''));

            if (!isNaN(numA)) sumQtyA += numA;
            if (!isNaN(numB)) sumQtyB += numB;
        }
    }

    let qtyHtml = "";
    if (hasQtyCol) {
        let matchColor = (sumQtyA === sumQtyB) ? "#10b981" : "#ef4444";
        qtyHtml = `
            <div class="big-stat"><div class="bs-val" style="color:#8b5cf6">${sumQtyA.toLocaleString()}</div><div class="bs-lbl">Form Qty</div></div>
            <div class="big-stat"><div class="bs-val" style="color:${matchColor}">${sumQtyB.toLocaleString()}</div><div class="bs-lbl">CPQ Qty</div></div>
        `;

        // Block with a modal when the quantity totals don't match.
        if (sumQtyA !== sumQtyB) {
            // Delay so the dashboard finishes painting before the modal appears.
            setTimeout(() => {
                showModal("Critical Quantity Mismatch", `The total Form Quantity (<b>${sumQtyA.toLocaleString()}</b>) does not match the CPQ Quantity (<b>${sumQtyB.toLocaleString()}</b>).<br><br>You must fix this discrepancy before proceeding!`, "error");
            }, 300);
        }
    }

    // Render the top summary stat cards.
    document.getElementById('globalStats').innerHTML = `
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${maxRows}</div><div class="bs-lbl">Rows</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#2563eb">${p.mapping.length}</div><div class="bs-lbl">Fields</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#10b981">${stats.matches}</div><div class="bs-lbl">Matches</div></div>
        <div class="big-stat"><div class="bs-val" style="color:#ef4444">${stats.mismatches}</div><div class="bs-lbl">Mismatches</div></div>
        ${qtyHtml}
    `; 
    
    // Render the column-health bar (active/mapped/unmapped counts per side).
    let colHealthBanner = document.getElementById('colHealthBanner');
    if (!colHealthBanner) {
        colHealthBanner = document.createElement('div');
        colHealthBanner.id = 'colHealthBanner';
        const statsNode = document.getElementById('globalStats');
        statsNode.parentNode.insertBefore(colHealthBanner, statsNode.nextSibling);
    }
    colHealthBanner.innerHTML = `
        <div style="display:flex; justify-content:space-between; background: #f8fafc; border: 1px solid #cbd5e1; padding: 15px 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); align-items: center;">
            <div style="text-align: center; flex: 1; border-right: 1px solid #e2e8f0;"><div style="font-size: 20px; font-weight: 800; color: #3b82f6;">${activeFormCols}</div><div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Active Form</div></div>
            <div style="text-align: center; flex: 1; border-right: 1px solid #e2e8f0;"><div style="font-size: 20px; font-weight: 800; color: #10b981;">${activeCPQCols}</div><div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Active CPQ</div></div>
            <div style="text-align: center; flex: 1; border-right: 1px solid #e2e8f0;"><div style="font-size: 20px; font-weight: 800; color: #8b5cf6;">${totalMapped}</div><div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Mapped</div></div>
            <div style="text-align: center; flex: 1; border-right: 1px solid #e2e8f0;"><div style="display: inline-block; padding: 2px 12px; border-radius: 6px; font-size: 20px; font-weight: 800; background: ${unmappedFormCols > 0 ? '#fee2e2' : 'transparent'}; color: ${unmappedFormCols > 0 ? '#b91c1c' : '#64748b'}; border: 1px solid ${unmappedFormCols > 0 ? '#fca5a5' : 'transparent'};">${unmappedFormCols}</div><div style="font-size: 10px; font-weight: 700; color: ${unmappedFormCols > 0 ? '#ef4444' : '#64748b'}; text-transform: uppercase; margin-top: 4px;">Unmapped Form</div></div>
            <div style="text-align: center; flex: 1;"><div style="display: inline-block; padding: 2px 12px; border-radius: 6px; font-size: 20px; font-weight: 800; background: ${unmappedCPQCols > 0 ? '#fee2e2' : 'transparent'}; color: ${unmappedCPQCols > 0 ? '#b91c1c' : '#64748b'}; border: 1px solid ${unmappedCPQCols > 0 ? '#fca5a5' : 'transparent'};">${unmappedCPQCols}</div><div style="font-size: 10px; font-weight: 700; color: ${unmappedCPQCols > 0 ? '#ef4444' : '#64748b'}; text-transform: uppercase; margin-top: 4px;">Unmapped CPQ</div></div>
        </div>`;
    

    renderResultTables(maxRows, doTrim); 
}

// Renders the two side-by-side result tables, colouring each cell by match status
// and flagging populated-but-unmapped columns.
function renderResultTables(maxRows, doTrim) {
    const p = projects[activeProjectIdx];
    const tA = document.getElementById('renderTableA');
    const tB = document.getElementById('renderTableB');

    // No data for this Set -> wipe the tables so a previous Set's rows don't linger.
    if (!p.dataA || !p.dataB) {
        if (tA) tA.innerHTML = "";
        if (tB) tB.innerHTML = "";
        return;
    }

    let hA = "<thead><tr><th>#</th>" + p.dataA.headers.map(h => `<th>${h}</th>`).join('') + "</tr></thead><tbody>";
    let hB = "<thead><tr><th>#</th>" + p.dataB.headers.map(h=>`<th>${h}</th>`).join('') + "</tr></thead><tbody>"; 
    
    let bA = ""; let bB = ""; 
    
    const mapLookup = {}; p.mapping.forEach(m => mapLookup[m.idxB] = m); 
    const reverseLookup = {}; 
    p.mapping.forEach(m => { 
        if (m.targetType === 'source') { 
            if (!reverseLookup[m.targetVal]) reverseLookup[m.targetVal] = []; 
            reverseLookup[m.targetVal].push(m); 
        } 
    }); 

    const currIdxB = p.dataB.headers.findIndex(h => /currency/i.test(h));

    for (let i = 0; i < maxRows; i++) { 
        let rA = `<td>${i+1}</td>`;
        let hasRowErrorA = false;

        // Form (Table 1) cells.
        for (let cA = 0; cA < p.dataA.headers.length; cA++) {
            let rawA = (p.dataA.body[i]?.[cA] || "").toString();
            let displayA = rawA; 
            let cls = ""; 
            let inlineStyle = ""; 

            if (reverseLookup[cA]) { 
                let allMatch = true;
                let comparedAgainst = ""; 
                let lastType = 'text';
                
                reverseLookup[cA].forEach(map => { 
                    let rawB = (p.dataB.body[i]?.[map.idxB] || "").toString(); 
                    comparedAgainst = rawB; 
                    let res = checkEqualAdvanced(rawA, rawB, map.name, doTrim);
                    if (!res.match) { allMatch = false; lastType = res.type; }
                }); 
                
                if (allMatch) { 
                    cls = "match"; 
                } else {
                    cls = lastType === 'qty' ? 'diff-qty' : (lastType === 'price' ? 'diff-price' : 'diff-text');
                    if (lastType === 'text') displayA = getVisualDiff(rawA, comparedAgainst);
                    hasRowErrorA = true; 
                }
            } 
            else {
                // Highlight Form columns that hold data but were never mapped.
                if (rawA.trim() !== "") {
                    cls = "unmapped-data-warning";
                    inlineStyle = "background-color: #e0f2fe; color: #0284c7; font-weight: bold; border: 2px dashed #7dd3fc; box-shadow: inset 0 0 5px rgba(0,0,0,0.05);";
                    displayA = `<i class="fas fa-exclamation-circle" style="color:#0ea5e9; margin-right:4px;" title="Warning: Data left unmapped"></i> ${displayA}`;
                }
            }

            rA += `<td class="${cls}" style="${inlineStyle}">${displayA}</td>`; 
        } 
        
        let rB = `<td>${i+1}</td>`;
        let hasRowErrorB = false;

        // CPQ (Table 2) cells.
        p.dataB.headers.forEach((_, colIdx) => {
            let rawB = (p.dataB.body[i]?.[colIdx] || "").toString();
            let displayB = rawB; 
            let cls = ""; 
            let inlineStyle = "";
            
            let isPriceB = /price|retail/i.test(p.dataB.headers[colIdx]);
            let numB = parseFloat(rawB.replace(/[^0-9.-]/g, ''));
            let currMissingB = isPriceB && !isNaN(numB) && numB > 0 && (currIdxB === -1 || String(p.dataB.body[i]?.[currIdxB]).trim() === "");

            if (mapLookup.hasOwnProperty(colIdx)) { 
                let map = mapLookup[colIdx];
                let rawA = map.targetType === 'matrix' ? (p.matrix.find(m => m.key === map.targetVal)?.val || "") : (p.dataA.body[i]?.[map.targetVal] || "").toString(); 
                
                let result = checkEqualAdvanced(rawA, rawB, map.name, doTrim);
                
                if (result.match) { 
                    cls = "match"; 
                } else {
                    cls = result.type === 'qty' ? 'diff-qty' : (result.type === 'price' ? 'diff-price' : 'diff-text');
                    if (result.type === 'text') displayB = getVisualDiff(rawB, rawA);
                    hasRowErrorB = true; 
                }
            } else {
                // Flag populated-but-unmapped CPQ columns (ignoring known filler columns).
                const headName = p.dataB.headers[colIdx].toLowerCase();
                if (rawB.trim() !== "" && headName !== "adjusted quantity" && headName !== "layout1") {
                    cls = "unmapped-data-warning";
                    inlineStyle = "background-color: #e0f2fe; color: #0284c7; font-weight: bold; border: 2px dashed #7dd3fc; box-shadow: inset 0 0 5px rgba(0,0,0,0.05);";
                    displayB = `<i class="fas fa-exclamation-circle" style="color:#0ea5e9; margin-right:4px;" title="Warning: Data left unmapped"></i> ${displayB}`;
                }
            }

            if (currMissingB) {
                cls = "diff-price";
                displayB = `${displayB} <span class="diff-badge" style="background:#ef4444; color:white;">Missing Currency</span>`;
                inlineStyle = ""; // drop the unmapped-blue box; the currency error wins
                hasRowErrorB = true;
            }
            
            rB += `<td class="${cls}" style="${inlineStyle}">${displayB}</td>`; 
        }); 

        const isRowError = hasRowErrorA || hasRowErrorB;
        const rowClass = isRowError ? "issue-row" : "perfect-row";
        
        bA += `<tr class="${rowClass}">${rA}</tr>`; 
        bB += `<tr id="rowB-${i}" class="${rowClass}">${rB}</tr>`; 
    } 
    
    tA.innerHTML = hA + bA + "</tbody>"; 
    tB.innerHTML = hB + bB + "</tbody>"; 
    
    toggleMismatchView();
}

// ==========================================
// HELPER UTILITIES
// ==========================================

// Serializes a 2-D array to tab/newline text, flattening line breaks inside cells.
function arrayToTSV(data) {
    if (!data) return "";
    return data.map(row =>
        row.map(cell => (cell == null ? "" : String(cell).replace(/[\r\n]+/g, " ").trim()))
            .join("\t")
    ).join("\n");
}


// Returns the index of the first row (within the top 60) that has a quantity column.
function findHeaderRowIndex(data) {
    const qtyRegex = /qty|quantity/i;
    for (let i = 0; i < Math.min(data.length, 60); i++) {
        if (data[i] && data[i].some(cell => cell && qtyRegex.test(String(cell).trim()))) {
            return i;
        }
    }
    return -1;
}


// ==========================================
// MATRIX RULES (key: value list above the table)
// ==========================================

// Handles edits to the matrix textarea: re-parses the rules, refreshes the UI, and
// propagates any universal ("*"-prefixed) rules to every other Set.
function handleMatrixInput(textarea, index = activeProjectIdx) {
    const p = projects[index];
    if (!p) return;

    p.rawMatrix = textarea.value;
    p.matrix = parseMatrixString(p.rawMatrix);

    // Rules whose key starts with "*" apply to all Sets.
    const universalRules = p.matrix.filter(m => m.key.startsWith('*'));

    if (universalRules.length > 0) {
        projects.forEach((proj, idx) => {
            if (idx !== index && proj.status !== 'empty') {
                let projRawLines = (proj.rawMatrix || "").split('\n').filter(l => l.trim() !== "");
                let changed = false;

                universalRules.forEach(uRule => {
                    // Skip rules the Set already has, so we don't clone endlessly.
                    if (!proj.matrix.some(m => m.key === uRule.key)) {
                        projRawLines.push(`${uRule.key}: ${uRule.val}`);
                        changed = true;
                    }
                });
                
                if (changed) {
                    proj.rawMatrix = projRawLines.join('\n');
                    proj.matrix = parseMatrixString(proj.rawMatrix);
                }
            }
        });
    }

    const listContainer = document.getElementById('matrixList');
    if (listContainer) {
        listContainer.innerHTML = ""; 
        p.matrix.forEach(m => addMatrixRow(m.key, m.val)); 
    }
}

// Paste handler for the matrix box: strips common boilerplate phrases before inserting.
function handleMatrixPaste(e, index = activeProjectIdx) {
    e.preventDefault();
    let paste = (e.clipboardData || window.clipboardData).getData('text');

    // Remove noisy vendor phrasing from pasted rules.
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
    textarea.value = newText;

    handleMatrixInput(textarea, index);
}

// Returns HTML for `mainText` with the words that differ from `compareText` wrapped
// in a highlight span (used to visualize text mismatches).
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

// Cleans parsed data in place: drops blank/placeholder header columns, tags each row
// with a stable id, and moves empty / repeated-header / zero-quantity rows into a
// hidden bucket (so they can be restored later).
function autoCleanData(data) {
    if (!data || !data.body) return;

    // Rows that fail the checks below go here, not deleted.
    data.hiddenRows = [];

    // Drop columns whose header is empty or just "#".
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

    // Give each row a hidden original index for the restore feature.
    data.body.forEach((row, i) => {
        if (!row.hasOwnProperty('_originalIdx')) {
            Object.defineProperty(row, '_originalIdx', {
                value: i,
                writable: true,
                enumerable: false
            });
        }
    });

    // Find the quantity columns; if there are none, keep every row.
    const qtyRegex = /qty|quantity|shipped|billed|units|pcs/i;
    const qtyIndices = [];
    data.headers.forEach((h, i) => {
        if (h && qtyRegex.test(h.toString().toLowerCase())) qtyIndices.push(i);
    });

    if (qtyIndices.length === 0) return;

    // Keep only rows that are non-empty, aren't a repeated header, and have qty > 0.
    const activeRows = [];

    data.body.forEach(row => {
        if (!row) return;
        const rowStr = row.join(" ").toLowerCase();

        const isEmpty = !row.some(cell => cell && cell.toString().trim() !== "");
        const isHeaderRepeat = rowStr.includes("upc") && rowStr.includes("style");

        let hasValidQty = false;
        for (let idx of qtyIndices) {
            let valStr = String(row[idx] || "").toLowerCase().replace(/[, \s]/g, '').replace(/pcs/g, '');
            let numVal = parseFloat(valStr);
            if (!isNaN(numVal) && numVal > 0) {
                hasValidQty = true;
                break;
            }
        }

        if (isEmpty || isHeaderRepeat || !hasValidQty) {
            data.hiddenRows.push({
                data: row,
                restoreIdx: row._originalIdx
            });
        } else {
            activeRows.push(row);
        }
    });

    data.body = activeRows;
}

// Parses tab-separated text (with quoted-cell support) into { headers, body }.
// `mode` "2rows" merges the first two rows into combined header names.
function parseExcelData(raw, mode) {
    if (!raw || !raw.trim()) return null;

    let text = raw.trim().replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let rows = [], currentRow = [], currentCell = "", insideQuote = false;

    for (let i = 0; i < text.length; i++) {
        let char = text[i], nextChar = text[i + 1];
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

    // Drop a leading "SKU Information" banner row if present.
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

// Normalizes a raw header: strips quotes, "(max N chars)" hints, bracketed notes,
// underscores and stray symbols, collapsing whitespace.
function cleanHeader(text) {
    if (!text) return "";

    let clean = text.trim();

    clean = clean
        .replace(/^"|"$/g, '')
        // Match "characters" before "chars" so nothing like "acters" is left behind.
        .replace(/\s*\(?max(imum)?\s*:?\s*\d+\s*(characters?|chars?|digits?)\)?/gi, "")
        .replace(/_/g, " ")
        .replace(/[\r\n]+/g, " ")
        .replace(/\s*\([^\)]*\)/g, "")  // remove ( ... )
        .replace(/\s*\[[^\]]*\]/g, "")  // remove [ ... ]
        .replace(/\s*\{[^\}]*\}/g, "")  // remove { ... }
        .replace(/[\*\^]/g, "")
        .trim();

    clean = clean.replace(/\s{2,}/g, " ");

    if (clean.includes("SIZEStyle")) return "Style";

    return clean;
}



// Shows a brief toast notification that auto-hides after 3 seconds.
function showToast(msg) {
    const t = document.getElementById('toast');
    t.querySelector('span').innerHTML = msg || "Done";
    t.classList.remove('hidden');
    setTimeout(() => { t.classList.add('hidden'); }, 3000);
}

// ==========================================
// DRAG-AND-DROP FILE UPLOAD
// ==========================================

document.addEventListener('DOMContentLoaded', () => {
    setupDragDrop('cardSource', 'A');
    setupDragDrop('cardTarget', 'B');
});

// Wires drag-and-drop on an upload card so dropped files go to the Form or CPQ handler.
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


// ==========================================
// HIDDEN / IGNORED ROWS (Step 2)
// ==========================================

// Toggles whether the ignored rows are shown for a side, then re-renders.
function toggleHiddenTable(elementId, btn) {
    const side = elementId.includes('A') ? 'A' : 'B';
    const p = projects[activeProjectIdx];

    if (!p.uiState) p.uiState = { showHiddenA: false, showHiddenB: false };

    if (side === 'A') {
        p.uiState.showHiddenA = !p.uiState.showHiddenA;
    } else {
        p.uiState.showHiddenB = !p.uiState.showHiddenB;
    }

    renderPreviewTables();
}

// Moves one row from the hidden bucket back into the active data.
function restoreHiddenRow(side, index) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;

    if (!dataObj || !dataObj.hiddenRows || !dataObj.hiddenRows[index]) return;

    const hiddenItem = dataObj.hiddenRows[index];
    const rowData = hiddenItem.data;
    const originalId = hiddenItem.restoreIdx;

    // Re-insert before the first active row with a higher original id, so it lands
    // back in its original position.
    let insertAt = dataObj.body.findIndex(r => (r._originalIdx !== undefined ? r._originalIdx : 999999) > originalId);
    if (insertAt === -1) {
        insertAt = dataObj.body.length;
    }

    dataObj.body.splice(insertAt, 0, rowData);
    dataObj.hiddenRows.splice(index, 1);
    renderPreviewTables();
}

// Writes an edited cell back to the model. Editing a quantity to 0/blank auto-hides
// the row; entering a valid quantity in a hidden row auto-restores it.
function updateCell(side, type, rowIndex, colIndex, cellElement) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    let newVal = cellElement.innerText.replace(/[\r\n]+/g, " ").trim();

    // Update the data model.
    if (type === 'active') {
        if (dataObj && dataObj.body && dataObj.body[rowIndex]) {
            dataObj.body[rowIndex][colIndex] = newVal;
        }
    } else if (type === 'hidden') {
        if (dataObj && dataObj.hiddenRows && dataObj.hiddenRows[rowIndex]) {
            dataObj.hiddenRows[rowIndex].data[colIndex] = newVal;
        }
    }

    // If this is a quantity column, move the row between active/hidden as needed.
    const headerName = dataObj.headers[colIndex] || "";
    const isQtyColumn = /^(qty|quantity|total\s*qty|total\s*quantity|units|pcs|bill|ship)$/i.test(headerName.trim());

    if (isQtyColumn) {
        const cleanVal = newVal.replace(/,/g, '');
        const numVal = parseFloat(cleanVal);
        const isValid = !isNaN(numVal) && numVal > 0 && newVal !== "";

        if (type === 'active' && !isValid) {
            // Quantity cleared/zeroed on an active row -> hide it.
            const row = dataObj.body[rowIndex];
            if (!dataObj.hiddenRows) dataObj.hiddenRows = [];
            dataObj.hiddenRows.push({
                data: row,
                restoreIdx: row._originalIdx ?? 999999
            });
            dataObj.body.splice(rowIndex, 1);
            renderPreviewTables();

        } else if (type === 'hidden' && isValid) {
            // Valid quantity entered on a hidden row -> bring it back.
            const hiddenItem = dataObj.hiddenRows[rowIndex];
            const row = hiddenItem.data;
            dataObj.body.push(row);
            dataObj.hiddenRows.splice(rowIndex, 1);
            renderPreviewTables();
        }
    }
}

// Saves an edit to a hidden row's cell (no active/hidden move).
function updateHiddenCell(side, rowIndex, colIndex, cellElement) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;

    if (dataObj && dataObj.hiddenRows && dataObj.hiddenRows[rowIndex]) {
        let newVal = cellElement.innerText.replace(/[\r\n]+/g, " ").trim();
        dataObj.hiddenRows[rowIndex].data[colIndex] = newVal;
    }
}

// ==========================================
// MANUAL TABLE SELECTOR (pick the table by dragging cells)
// ==========================================

// Column bounds of the current manual-selection rectangle (rows tracked separately).
let selectionStartCol = -1;
let selectionEndCol = -1;

// Reopens the manual selector for the active Set using its own source workbook/sheet.
function reselectCurrentTable() {
    const p = projects[activeProjectIdx];

    // Prefer the workbook saved on this specific Set.
    const workbook = p ? p.sourceWorkbook : null;
    const filename = p ? p.fileName : "";

    if (!workbook) {
        // Fallback: If no specific workbook is saved, try the global last uploaded one
        if (lastUploadedWorkbook) {
            openManualMapper(lastUploadedWorkbook, lastUploadedFilename, currentSheetName);
        } else {
            alert("No file is currently loaded in memory for this set. Please re-upload the file.");
        }
        return;
    }

    const availableSheets = workbook.SheetNames;
    let targetSheet = availableSheets[0]; // Default to first sheet

    // 2. Select the EXACT sheet that this Set represents
    if (p) {
        if (p.originalSheetName && availableSheets.includes(p.originalSheetName)) {
            targetSheet = p.originalSheetName;
        } else if (availableSheets.includes(p.name)) {
            targetSheet = p.name;
        }
    }

    // 3. Open the mapper with the correct specific file and sheet
    openManualMapper(workbook, filename, targetSheet);
}


// ==========================================
// MANUAL-SELECTION QUEUE
// When several uploaded files all fail auto-detection, walk the user through them
// one-by-one instead of only prompting for a single file.
// ==========================================

// Begins stepping through the queued files that need manual table selection.
function startManualQueue() {
    if (!manualQueue || manualQueue.length === 0) return;
    manualBatchActive = true;
    manualQueueTotal = manualQueue.length;
    processNextManual();
}

// Opens the mapper for the next queued file, or finishes the batch when empty.
function processNextManual() {
    if (!manualQueue || manualQueue.length === 0) {
        manualBatchActive = false;
        renderTopBar();

        // If this batch also had CPQ files waiting, inject them now across every
        // form set created this batch (auto-detected + manually selected).
        if (window.pendingBatchCPQs && window.pendingBatchCPQs.length > 0) {
            const created = projects.length - batchStartCount;
            const cpqs = window.pendingBatchCPQs;
            window.pendingBatchCPQs = null;
            manualQueueTotal = 0;
            if (created > 0) { processBatchCPQs(cpqs, created); return; }
        }

        showModal("Manual Selection Complete",
            `Finished reviewing all <strong>${manualQueueTotal}</strong> file(s) that needed manual selection.`,
            "success");
        manualQueueTotal = 0;
        return;
    }

    const item = manualQueue[0];
    openManualMapper(item.workbook, item.filename);
    updateManualQueueBanner();
}

// Called whenever ONE manual-selection session finishes (import or skip)
function advanceManualQueue() {
    if (!manualBatchActive) return;
    manualQueue.shift();                 // drop the file we just handled
    setTimeout(processNextManual, 150);  // let the current modal close first
}

// Show "File X of Y — filename" progress in the modal header
function updateManualQueueBanner() {
    const subtitle = document.getElementById('manualSubtitle');
    if (!subtitle) return;
    if (manualBatchActive && manualQueueTotal > 1) {
        const done = manualQueueTotal - manualQueue.length;
        const current = manualQueue[0];
        subtitle.innerHTML = `<span style="background:#f59e0b; color:#fff; padding:1px 8px; border-radius:10px; font-weight:700;">File ${done + 1} of ${manualQueueTotal}</span> &nbsp;<b>${current ? current.filename : ''}</b> — select the table, or Skip to move on.`;
    } else {
        subtitle.innerText = "Automation failed. Please select the Header Row manually.";
    }
}


// Opens the manual table-selection modal for a workbook: populates the sheet
// dropdown (visible sheets only) and renders the chosen sheet's raw grid.
function openManualMapper(workbook, filename, targetSheetName = null) {
    manualWorkbook = workbook;
    manualFilename = filename;

    const select = document.getElementById('manualSheetSelect');
    select.innerHTML = "";
    if (workbook.SheetNames.length === 0) return alert("Empty Excel");

    let firstVisibleSheet = null;

    workbook.SheetNames.forEach(name => {
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

    // Prefer the requested sheet, else the first visible one.
    const sheetToLoad = targetSheetName || firstVisibleSheet || workbook.SheetNames[0];
    currentSheetName = sheetToLoad;
    renderRawSheet(sheetToLoad);
}

// Renders a sheet as a raw, selectable grid in the manual selector, with the
// smart/manual mode toolbar on top.
function renderRawSheet(sheetName) {
    currentSheetName = sheetName;

    // Clear any previous selection.
    selectionStartRow = -1; selectionEndRow = -1;
    selectionStartCol = -1; selectionEndCol = -1;
    if (typeof updateManualButtonState === "function") updateManualButtonState();

    const sheet = manualWorkbook.Sheets[sheetName];
    const table = document.getElementById('manualRawTable');
    const container = table.parentElement;

    // Build the sticky toolbar (once).
    let toolbar = document.getElementById('manualToolbar');
    if (!toolbar) {
        toolbar = document.createElement('div');
        toolbar.id = 'manualToolbar';
        toolbar.style.cssText = "display:flex; justify-content:space-between; align-items:center; background:#f8fafc; padding:12px 20px; border-bottom:1px solid #e2e8f0; width:100%; box-sizing:border-box; z-index:100; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);";
        
        const modalBody = container.parentElement;
        modalBody.insertBefore(toolbar, container);
    }
    
    toolbar.innerHTML = `
        <div style="display:flex; gap:15px; align-items:center;">
            <label style="font-size:12px; font-weight:bold; color:#334155; display:flex; align-items:center; gap:5px; cursor:pointer; background:white; padding:6px 12px; border:1px solid #cbd5e1; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05);">
                <input type="radio" name="manMode" value="smart" checked onchange="window.manualSelectMode='smart'"> ⚡ Smart Auto-Select
            </label>
            <label style="font-size:12px; font-weight:bold; color:#334155; display:flex; align-items:center; gap:5px; cursor:pointer; background:white; padding:6px 12px; border:1px solid #cbd5e1; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05);">
                <input type="radio" name="manMode" value="manual" onchange="window.manualSelectMode='manual'"> 🖐️ Manual Drag
            </label>
        </div>
        <div>
            <button onclick="addSelectionToMatrix()" style="background:#8b5cf6; color:white; border:none; padding:8px 16px; border-radius:6px; font-size:12px; font-weight:bold; cursor:pointer; box-shadow:0 2px 4px rgba(139,92,246,0.3); transition:all 0.2s;"><i class="fas fa-plus"></i> Add to Matrix Rules</button>
        </div>
    `;
    
    window.manualSelectMode = 'smart';

    // Rebuild the grid (capped at 300 rows / 40 cols) from the sheet's cells.
    table.innerHTML = "";
    table.setAttribute('tabindex', '0');
    table.style.outline = 'none';

    // New sheet -> reset undo/redo history and clear any stale selection.
    excelHistory = [];
    excelRedo = [];
    excelSelStart = null;
    excelSelEnd = null;

    if (!sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const maxRows = Math.min(range.e.r, 300);
    const maxCols = Math.min(range.e.c, 40);

    for(let r = range.s.r; r <= maxRows; r++) {
        const tr = document.createElement('tr');

        // Sticky row-number cell on the left.
        let html = `<td class="excel-idx" style="background:#f1f5f9; text-align:center; color:#888; font-size:10px; user-select:none; position:sticky; left:0; z-index:5;">${r+1}</td>`;
        
        for(let c = range.s.c; c <= maxCols; c++) {
            const cellAddr = XLSX.utils.encode_cell({r, c});
            const val = sheet[cellAddr] ? XLSX.utils.format_cell(sheet[cellAddr]) : "";
            
            html += `<td id="cell-${r}-${c}" 
                        data-r="${r}" data-c="${c}"
                        style="padding:6px; border:1px solid #e2e8f0; cursor:cell; min-width:50px; overflow:hidden; white-space:nowrap; max-width:200px; user-select: none;">
                        ${val}
                      </td>`;
        }
        tr.innerHTML = html;
        table.appendChild(tr);
    }

    setTimeout(() => { enableExcelFeatures('manualRawTable'); }, 100);
}


// Turns the highlighted cells in the manual selector into "Key: Value" matrix rules,
// treating the first cell of each row as the key and the last as the value.
function addSelectionToMatrix() {
    if (!excelSelStart || !excelSelEnd) return alert("Please highlight cells first.");

    const rMin = Math.min(excelSelStart.r, excelSelEnd.r);
    const rMax = Math.max(excelSelStart.r, excelSelEnd.r);
    const cMin = Math.min(excelSelStart.c, excelSelEnd.c);
    const cMax = Math.max(excelSelStart.c, excelSelEnd.c);

    const table = document.getElementById('manualRawTable');
    let newRulesText = "";
    let addedCount = 0;

    for (let r = rMin; r <= rMax; r++) {
        let rowValues = [];

        // Collect the non-empty text across the highlighted columns of this row.
        for (let c = cMin; c <= cMax; c++) {
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell) {
                const text = cell.innerText.trim().replace(/[\r\n]+/g, " ");
                if (text) rowValues.push(text);
            }
        }

        if (rowValues.length === 0) continue;

        // Drop known filler words so they don't get mistaken for the value.
        rowValues = rowValues.filter((val, index) => {
            if (index === 0) return true; // always keep the key
            if (val === "Department#:" || val === "Department:") return false;
            return true;
        });

        // Build the rule: first cell is the key, last is the value.
        if (rowValues.length >= 2) {
            let k = rowValues[0].replace(/:$/, ""); // strip a trailing colon
            let v = rowValues[rowValues.length - 1];

            newRulesText += `${k}: ${v}\n`;
            addedCount++;
        } else if (rowValues.length === 1) {
            newRulesText += `${rowValues[0]}\n`;
            addedCount++;
        }
    }

    if (addedCount > 0) {
        const p = projects[activeProjectIdx];
        const textarea = document.getElementById('matrixRawInput');

        // Append the new rules to the matrix textarea.
        let currentText = textarea.value.trim();
        if (currentText) {
            textarea.value = currentText + "\n" + newRulesText.trim();
        } else {
            textarea.value = newRulesText.trim();
        }

        // Make sure the matrix panel is open.
        p.showMatrix = true;
        const matSec = document.getElementById('matrixSection');
        const btn = document.getElementById('btnToggleMatrix');
        if (matSec && btn) {
            matSec.style.display = 'block';
            btn.innerHTML = `<i class="fas fa-minus-circle"></i> Hide Matrix Rules`;
        }

        // Re-run the matrix parser as if the rules were typed.
        if (typeof handleMatrixInput === 'function') {
            handleMatrixInput(textarea, activeProjectIdx);
        }

        showToast(`Added ${addedCount} rule(s) to text box!`);
        for (let r = rMin; r <= rMax; r++) {
            for (let c = cMin; c <= cMax; c++) {
                const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                if (cell) {
                    cell.style.transition = "background 0.4s";
                    cell.style.backgroundColor = "#dcfce7";
                    setTimeout(() => { cell.style.backgroundColor = ""; }, 600);
                }
            }
        }
    }
}


// Begins a drag selection in the manual grid and starts auto-scroll tracking.
function handleCellMouseDown(el) {
    isDragging = true;
    selectionStartRow = parseInt(el.getAttribute('data-r'));
    selectionStartCol = parseInt(el.getAttribute('data-c'));

    selectionEndRow = selectionStartRow;
    selectionEndCol = selectionStartCol;

    highlightSelection();
    updateManualButtonState();

    document.addEventListener('mousemove', handleDragAutoScroll);
}

// Extends the selection to the cell under the cursor while dragging.
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

// Repaints the manual grid to show the current selection rectangle (header row in blue).
function highlightSelection() {
    // Clear previous highlighting and reset borders.
    const cells = document.getElementById('manualRawTable').querySelectorAll('td');
    cells.forEach(el => {
        if (el.classList.contains('excel-idx')) return;
        el.style.backgroundColor = "transparent";
        el.style.color = "inherit";
        el.style.borderColor = "#e2e8f0";
    });

    if (selectionStartRow === -1) return;

    const rMin = Math.min(selectionStartRow, selectionEndRow);
    const rMax = Math.max(selectionStartRow, selectionEndRow);
    const cMin = Math.min(selectionStartCol, selectionEndCol);
    const cMax = Math.max(selectionStartCol, selectionEndCol);

    for (let r = rMin; r <= rMax; r++) {
        for (let c = cMin; c <= cMax; c++) {
            const cell = document.getElementById(`cell-${r}-${c}`);
            if (cell) {
                // Match border to background so there's no white gap between cells.
                cell.style.backgroundColor = "#e0f2fe";
                cell.style.borderColor = "#e0f2fe";

                if (r === rMin) {
                    // Emphasize the header row.
                    cell.style.backgroundColor = "#2563eb";
                    cell.style.borderColor = "#2563eb";
                    cell.style.color = "white";
                }
            }
        }
    }
}

// Enables/disables the Confirm button and updates the instruction text based on
// whether a selection exists.
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


// Imports the selected rectangle from the manual grid as the Set's Form data.
// In batch mode each file becomes a new Set; otherwise the user is asked whether
// to replace the active Set or create a new one.
function confirmManualImport() {
    if (selectionStartRow === -1) return;

    try {
        const sheet = manualWorkbook.Sheets[currentSheetName];
        const rawData = extractSmartExcelData(sheet);

        // Slice out the selected rows/columns.
        const rMin = Math.min(selectionStartRow, selectionEndRow);
        const rMax = Math.max(selectionStartRow, selectionEndRow);
        const cMin = Math.min(selectionStartCol, selectionEndCol);
        const cMax = Math.max(selectionStartCol, selectionEndCol);

        let slicedData = [];
        for (let r = rMin; r <= rMax; r++) {
            let newRow = [];
            const srcRow = rawData[r] || [];
            for (let c = cMin; c <= cMax; c++) {
                newRow.push(srcRow[c] || "");
            }
            slicedData.push(newRow);
        }

        if (slicedData.length < 2) {
            alert("Please select at least 2 rows (1 Header + 1 Data).");
            return;
        }

        const newRawData = arrayToTSV(slicedData);

        // Batch mode: each queued file becomes its own new Set (no Replace/Create prompt).
        if (manualBatchActive) {
            applyManualNew(newRawData);
            advanceManualQueue();
            return;
        }

        // Otherwise apply to the tab the user is editing (or the active one).
        let targetIdx = typeof editingProjectIndex !== 'undefined' && editingProjectIndex !== -1
            ? editingProjectIndex
            : activeProjectIdx;

        if (targetIdx !== -1 && projects[targetIdx]) {
            // Rename the tab to the selected sheet and ask replace-vs-new.
            projects[targetIdx].name = currentSheetName;
            showConflictModal(currentSheetName, targetIdx, newRawData);
        } else {
            // No active tab — just create a new Set.
            applyManualNew(newRawData);
        }

    } catch (err) {
        console.error(err);
        alert("Error importing: " + err.message);
    }
}

// Closes the manual selector. During a batch, "Skip / Close" advances to the next file.
function closeManualModal() {
    document.getElementById('manualSelectModal').style.display = 'none';
    manualWorkbook = null;
    manualFilename = "";
    advanceManualQueue();
}

// ==========================================
// MANUAL SELECTOR EDGE AUTO-SCROLL
// ==========================================
let scrollInterval = null;
let scrollVector = { x: 0, y: 0 };

// While dragging near an edge of the manual grid, scrolls the container that way.
function handleDragAutoScroll(e) {
    if (!isDragging) return;

    const table = document.getElementById('manualRawTable');
    if (!table) return;
    const container = table.parentElement; // scrollable wrapper
    if (!container) return;

    const rect = container.getBoundingClientRect();
    const buffer = 50; // px from the edge that triggers scrolling
    const speed = 20;  // px per tick

    let vx = 0;
    let vy = 0;

    if (e.clientX > rect.right - buffer) vx = speed;
    else if (e.clientX < rect.left + buffer) vx = -speed;

    if (e.clientY > rect.bottom - buffer) vy = speed;
    else if (e.clientY < rect.top + buffer) vy = -speed;

    scrollVector = { x: vx, y: vy };

    // Run a scroll interval only while near an edge.
    if (vx !== 0 || vy !== 0) {
        if (!scrollInterval) {
            scrollInterval = setInterval(() => {
                container.scrollLeft += scrollVector.x;
                container.scrollTop += scrollVector.y;
            }, 30);
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
// STEP 2: ADD ROW & PASTE
// ==========================================

// Appends a blank row (with a fresh original id) to a Step 2 table.
function addNewRow(side) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;

    const colCount = dataObj.headers.length;
    const newRow = new Array(colCount).fill("");

    // New id = one past the highest existing (active or hidden) id.
    let maxId = 0;
    if (dataObj.body.length > 0) maxId = Math.max(...dataObj.body.map(r => r._originalIdx || 0));
    if (dataObj.hiddenRows && dataObj.hiddenRows.length > 0) {
        const hiddenMax = Math.max(...dataObj.hiddenRows.map(h => h.restoreIdx || 0));
        maxId = Math.max(maxId, hiddenMax);
    }

    Object.defineProperty(newRow, '_originalIdx', { value: maxId + 1, writable: true, enumerable: false });

    dataObj.body.push(newRow);

    renderPreviewTables();
    showToast("Row Added");
}

// Multi-cell paste into a Step 2 table starting at the given cell.
function handlePaste(e, side, rowIndex, colIndex) {
    // Prevent the default single-cell paste.
    e.preventDefault();
    e.stopPropagation();

    // Read the clipboard grid.
    const clipboardData = (e.clipboardData || window.clipboardData).getData('text');
    if (!clipboardData) return;

    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;
    const rows = dataObj.body;

    // 3. Split into Rows and Columns (Excel uses \n for rows, \t for cols)
    const pasteRows = clipboardData.split(/\r\n|\n|\r/);

    let rowsUpdated = 0;

    pasteRows.forEach((pasteRow, i) => {
        if (!pasteRow && i === pasteRows.length - 1) return; // ignore trailing newline

        const targetRowIdx = rowIndex + i;
        if (targetRowIdx >= rows.length) return; // ran out of rows

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


// NOTE: addNewRow and handlePaste are (re)defined below; these later definitions are
// the ones actually used. handlePaste here also lets native paste through while a
// cell is being edited.
function addNewRow(side) {
    const p = projects[activeProjectIdx];
    const dataObj = (side === 'A') ? p.dataA : p.dataB;

    const colCount = dataObj.headers.length;
    const newRow = new Array(colCount).fill("");

    // New id = one past the highest existing (active or hidden) id.
    let maxId = 0;
    if (dataObj.body.length > 0) maxId = Math.max(...dataObj.body.map(r => r._originalIdx || 0));
    if (dataObj.hiddenRows && dataObj.hiddenRows.length > 0) {
        const hiddenMax = Math.max(...dataObj.hiddenRows.map(h => h.restoreIdx || 0));
        maxId = Math.max(maxId, hiddenMax);
    }

    Object.defineProperty(newRow, '_originalIdx', { value: maxId + 1, writable: true, enumerable: false });

    dataObj.body.push(newRow);
    renderPreviewTables();
    showToast("Row Added");
}

// Grid paste into a Step 2 table; leaves native paste alone while typing in a cell.
function handlePaste(e, side, rowIndex, colIndex) {
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
// MATRIX EXTRACTION HELPERS
// ==========================================

// Returns the first row index (top 60) that contains a quantity column.
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


// Scans the rows above the table (before `endRow`) for known "Key: Value" matrix
// fields and returns them as {key, val} pairs.
function extractMatrixData(data, endRow) {
    const matrix = [];
    // Keywords are upper-cased for case-insensitive matching.
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

            // Skip instruction/option header cells.
            if (upperCell.includes("INSTRUCTION")) continue;
            if (upperCell.includes("FEATURE OPTION")) continue;
            if (upperCell.includes("FEATURES OPTION")) continue;
            if (upperCell.includes("DRAW FROM BULK")) continue;

            const matchedKeyword = keywords.find(k => upperCell.includes(k));

            if (matchedKeyword) {
                let candidateVal = "";
                let foundVal = false;

                // Value is the next cell to the right...
                if (row[c + 1] && String(row[c + 1]).trim() !== "") {
                    candidateVal = String(row[c + 1]).trim();
                    foundVal = true;
                }
                // ...or two cells over, to jump a merged/empty cell.
                else if (row[c + 2] && String(row[c + 2]).trim() !== "") {
                    candidateVal = String(row[c + 2]).trim();
                    foundVal = true;
                }

                // Guard against the "value" actually being another label.
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

                if (!matrix.some(m => m.key === cleanKey)) {
                    matrix.push({ key: cleanKey, val: finalVal });
                }
            }
        }
    }
    return matrix;
}

// Parses the matrix textarea text into {key, val} rules, normalizing well-known keys.
function parseMatrixString(rawString) {
    if (!rawString) return [];
    const rows = [];
    const rawLines = rawString.split(/\n/);

    rawLines.forEach(line => {
        let cleanLine = line.trim();

        let k = "", v = "";

        // Split on the first colon before stripping anything from the key.
        if (cleanLine.includes(":")) {
            let idx = cleanLine.indexOf(":");
            k = cleanLine.substring(0, idx);
            v = cleanLine.substring(idx + 1);
        } else {
            k = cleanLine;
        }

        // Key: drop quotes/parentheses and collapse spaces. Value: trim only
        // (parentheses in values are meaningful).
        k = k.replace(/"/g, "").replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim();
        v = v ? v.trim() : "";

        // Canonicalize recognized keys.
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

// Parses the matrix box (colon/tab/equals separated) and rebuilds the editable rows.
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
            v = row.substring(idx + 1);
        } else if (row.includes("=")) {
            let idx = row.indexOf("=");
            k = row.substring(0, idx);
            v = row.substring(idx + 1);
        } else {
            k = row;
            v = "";
        }
        addMatrixRow(mapKey(cleanLabel(k || "")), normalizeValue(v || ""));
    });
}

// Appends one editable key/value input row to the matrix list.
function addMatrixRow(key = "", val = "") {
    const div = document.createElement('div'); div.className = 'matrix-row-item';
    div.innerHTML = `<input type="text" class="matrix-input m-key" placeholder="Field" value="${key}"><input type="text" class="matrix-input m-val" placeholder="Value" value="${val}"><button class="btn-x" onclick="this.parentElement.remove()">×</button>`;
    document.getElementById('matrixList').appendChild(div);
}

// Shows/hides the optional matrix-rules panel.
function toggleMatrix() {
    const sec = document.getElementById('matrixSection'), btn = document.getElementById('btnToggleMatrix');
    if (sec.style.display === 'none') { sec.style.display = 'block'; btn.innerHTML = `<i class="fas fa-minus-circle"></i> Hide Matrix Rules`; }
    else { sec.style.display = 'none'; btn.innerHTML = `<i class="fas fa-plus-circle"></i> Show Matrix Rules (Optional)`; }
}

// Reads the editable matrix rows back into {key, val} objects.
function getMatrixDataFromUI() {
    const rows = document.querySelectorAll('.matrix-row-item');
    const data = [];
    rows.forEach(r => {
        const k = r.querySelector('.m-key').value.trim();
        const v = r.querySelector('.m-val').value.trim();
        if (k) data.push({ key: k, val: v });
    });
    return data;
}

// Small string helpers for matrix keys/values.
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
// APPLYING A MANUAL IMPORT (replace vs. create)
// ===============================================

// Asks whether a manually selected sheet should replace the active tab or become
// a new one, then dispatches to applyManualReplace / applyManualNew.
function showConflictModal(name, existingIdx, newData) {
    const modal = document.getElementById('customModal');
    const titleEl = document.getElementById('modalTitle');
    const msgEl = document.getElementById('modalMsg');
    const iconEl = document.getElementById('modalIcon');

    const confirmBtn = document.getElementById('modalBtn');
    const cancelBtn = document.getElementById('modalCancelBtn');

    titleEl.innerText = "Sheet Already Exists";
    msgEl.innerHTML = `The sheet <b>"${name}"</b> is already loaded.<br>Do you want to REPLACE the existing tab or Create a NEW one?`;

    iconEl.className = 'fas fa-question-circle sa-icon-warn';
    iconEl.style.color = '#f59e0b';

    // Primary button = Replace. Clone it first to drop any stale click handlers.
    confirmBtn.innerText = "Replace Existing";
    confirmBtn.style.backgroundColor = '#f59e0b';
    const newConfirm = confirmBtn.cloneNode(true);
    confirmBtn.parentNode.replaceChild(newConfirm, confirmBtn);

    newConfirm.onclick = function () {
        applyManualReplace(existingIdx, newData);
        closeModal();
    };

    // Secondary button = Create New (also cloned to reset handlers).
    cancelBtn.style.display = 'inline-block';
    cancelBtn.innerText = "Create New Tab";
    cancelBtn.style.backgroundColor = '#2563eb';
    cancelBtn.style.color = 'white';

    const newCancel = cancelBtn.cloneNode(true);
    cancelBtn.parentNode.replaceChild(newCancel, cancelBtn);

    newCancel.onclick = function () {
        applyManualNew(newData);
        closeModal();
    };

    modal.classList.add('open');
}

// Replaces an existing Set's Form data with the manually imported table.
function applyManualReplace(idx, rawData) {
    const p = projects[idx];
    p.rawA = rawData;
    p.dataA = null;
    p.status = 'ready';
    p.step = 1;

    // skipSave=true so we don't overwrite the data we just imported.
    switchProject(idx, true);

    document.getElementById('manualSelectModal').style.display = 'none';
    showToast(`Replaced data in "${p.name}"`);
}

// Creates a new Set from the manually imported table (remembering its workbook).
function applyManualNew(rawData) {
    createSet(currentSheetName, manualWorkbook, currentSheetName, manualFilename);
    const newIdx = projects.length - 1;

    projects[newIdx].rawA = rawData;
    projects[newIdx].status = 'ready';
    projects[newIdx].uploadSessionId = Date.now();

    document.getElementById('manualSelectModal').style.display = 'none';
    renderTopBar();
    switchProject(newIdx, true);
    showToast(`Created "${currentSheetName}"`);
}


// ==========================================
// EXCEL-LIKE GRID ENGINE (manual selector: drag-select, copy/paste, undo)
// ==========================================
let excelSelStart = null;
let excelSelEnd = null;
let isExcelDragging = false;
let autoScrollTimer = null;
let lastMouseX = 0;
let lastMouseY = 0;

// Edit-Mode support: undo/redo history plus a hidden clipboard proxy.
let excelHistory = [];      // stack of table snapshots (for Undo)
let excelRedo = [];         // stack of table snapshots (for Redo)
let excelClipboardProxy = null; // hidden textarea that captures native copy/paste (works on file://)

// Wires up drag-select, auto-scroll, and clipboard/keyboard handling on the manual
// selector grid. In "smart" mode a click auto-selects the table downward; in "manual"
// mode the user drags the exact rectangle.
function enableExcelFeatures(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return;

    const container = table.parentElement; // scrollable wrapper
    container.style.position = 'relative';

    // Ensure the hidden clipboard proxy exists (powers Ctrl+C/V/Z in Edit Mode).
    ensureExcelProxy(table);

    // Mouse down: begin (or shift-extend) a selection.
    table.onmousedown = function (e) {
        // Let the browser handle text selection inside a cell that's open for editing.
        if (e.target.isContentEditable) return;
        // Ignore non-left clicks and form controls.
        if (e.button !== 0 || ['INPUT', 'BUTTON', 'SELECT'].includes(e.target.tagName)) return;

        const cell = e.target.closest('td');
        if (!cell || cell.classList.contains('excel-idx')) return;

        isExcelDragging = true;
        const startR = parseInt(cell.getAttribute('data-r'));
        const c = parseInt(cell.getAttribute('data-c'));

        const maxR = table.rows.length - 1;
        const maxC = (table.rows[0] ? table.rows[0].cells.length - 2 : 20);

        if (e.shiftKey && excelSelStart) {
            // Shift-click extends the current selection.
            excelSelEnd = { r: startR, c: c };
        } else {
            if (window.manualSelectMode === 'manual') {
                // Manual mode: start a single-cell selection to drag from.
                excelSelStart = { r: startR, c: c };
                excelSelEnd = { r: startR, c: c };
            } else {
                // Smart mode: auto-extend down to the end of the table, stopping at
                // blank gaps or footer/note rows.
                let endR = maxR;
                let emptyCount = 0;

                for (let i = startR + 1; i <= maxR; i++) {
                    const rowTr = table.rows[i];
                    if (!rowTr) continue;

                    let rowTextArray = [];
                    let hasData = false;
                    for (let col = 1; col < rowTr.cells.length; col++) { 
                        if(rowTr.cells[col].classList.contains('excel-idx')) continue;
                        
                        const cellText = rowTr.cells[col].innerText.trim().toLowerCase();
                        rowTextArray.push(cellText);
                        if (cellText) hasData = true;
                    }
                    const rowTextStr = rowTextArray.join(" ");

                    if (!hasData) {
                        emptyCount++;
                        if (emptyCount >= 5) { endR = i - 5; break; }
                        continue;
                    } else {
                        emptyCount = 0;
                    }

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
                        endR = i - 1; 
                        break;
                    }

                    const firstText = rowTextArray.find(t => t.length > 0);
                    if (firstText && (firstText === "total" || firstText.startsWith("total:") || firstText.startsWith("total qty"))) {
                        endR = i - 1;
                        break;
                    }
                }

                if (endR < startR) endR = startR; 
                excelSelStart = { r: startR, c: c }; 
                excelSelEnd = { r: endR, c: maxC };
            }
        }

        // Keep the older row/col globals in sync for code that still reads them.
        updateLegacyGlobals(excelSelStart.r, excelSelStart.c, excelSelEnd.r, excelSelEnd.c);

        highlightExcelRange(table);
        // Edit Mode: keep the proxy focused for Ctrl+C/V/Z; else focus the table.
        if (isEditMode) {
            syncExcelClipboard(table);
        } else {
            table.focus();
        }

        if (typeof startAutoScroll === 'function') startAutoScroll(container);
    };

    // Mouse move (on document, so dragging works even outside the table): extend
    // the selection to the cell under the cursor.
    document.addEventListener('mousemove', function (e) {
        if (!isExcelDragging) return;

        // Track cursor position for the auto-scroller.
        if (typeof lastMouseX !== 'undefined') lastMouseX = e.clientX;
        if (typeof lastMouseY !== 'undefined') lastMouseY = e.clientY;

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

    // Mouse up (on document): end the drag and re-sync the clipboard proxy.
    document.addEventListener('mouseup', function () {
        const wasDragging = isExcelDragging;
        isExcelDragging = false;
        if (typeof stopAutoScroll === 'function') stopAutoScroll();
        // Point the proxy at the final selection so Ctrl+C copies the right block.
        if (wasDragging && isEditMode && excelSelStart) {
            const t = document.getElementById('manualRawTable');
            if (t) syncExcelClipboard(t);
        }
    });

    // Keyboard: in Edit Mode delegate to the full handler; otherwise support
    // Ctrl+C / Ctrl+V / arrow-key navigation.
    table.onkeydown = async function (e) {
        if (!excelSelStart) return;

        if (isEditMode) { handleExcelProxyKey(e, table); return; }

        if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
            e.preventDefault();
            copyExcelData(table);
        }

        if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
            e.preventDefault();
            try {
                const text = await navigator.clipboard.readText();
                pasteExcelData(table, text);
            } catch (err) { }
        }

        if (e.key.startsWith('Arrow')) {
            e.preventDefault();
            moveExcelSelection(e.key, table);
        }
    };
}

// Auto-scrolls the grid container while the cursor is near an edge during a drag.
function startAutoScroll(container) {
    stopAutoScroll();

    autoScrollTimer = setInterval(() => {
        if (!isExcelDragging) return;

        const rect = container.getBoundingClientRect();
        const sensitivity = 50; // px from edge that triggers scrolling
        const speed = 15;       // px per tick

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

        // After scrolling, extend the selection to whatever cell is now under the cursor.
        if (scrolled) {
            const el = document.elementFromPoint(lastMouseX, lastMouseY);
            const cell = el ? el.closest('td') : null;
            if (cell) {
                const r = parseInt(cell.getAttribute('data-r'));
                const c = parseInt(cell.getAttribute('data-c'));
                if (!isNaN(r) && !isNaN(c)) {
                    excelSelEnd = { r, c };
                    if (typeof selectionEndRow !== 'undefined') {
                        selectionEndRow = r; selectionEndCol = c;
                    }
                    const table = cell.closest('table');
                    if (table) highlightExcelRange(table);
                }
            }
        }
    }, 30);
}

function stopAutoScroll() {
    if (autoScrollTimer) clearInterval(autoScrollTimer);
    autoScrollTimer = null;
}

// Mirrors the current selection into the older selectionStart/End globals and
// refreshes the Confirm button.
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

// Highlights the selected rectangle by toggling the 'excel-selected' class.
function highlightExcelRange(table) {
    table.querySelectorAll('.excel-selected').forEach(el => el.classList.remove('excel-selected'));

    if (!excelSelStart || !excelSelEnd) return;

    const r1 = Math.min(excelSelStart.r, excelSelEnd.r);
    const r2 = Math.max(excelSelStart.r, excelSelEnd.r);
    const c1 = Math.min(excelSelStart.c, excelSelEnd.c);
    const c2 = Math.max(excelSelStart.c, excelSelEnd.c);

    for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell) cell.classList.add('excel-selected');
        }
    }
}

// Serializes the selected cells to TSV (tabs between columns, newlines between rows).
function buildSelectionTSV(table) {
    if (!table || !excelSelStart || !excelSelEnd) return null;
    const r1 = Math.min(excelSelStart.r, excelSelEnd.r);
    const r2 = Math.max(excelSelStart.r, excelSelEnd.r);
    const c1 = Math.min(excelSelStart.c, excelSelEnd.c);
    const c2 = Math.max(excelSelStart.c, excelSelEnd.c);

    let lines = [];
    for (let r = r1; r <= r2; r++) {
        let cols = [];
        for (let c = c1; c <= c2; c++) {
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            cols.push(cell ? cell.innerText.replace(/\r?\n/g, ' ') : '');
        }
        lines.push(cols.join('\t'));
    }
    return lines.join('\n'); // tabs = columns, newlines = rows (real Excel/TSV format)
}

// Copies the selected cells to the clipboard (async API + proxy fallback).
function copyExcelData(table) {
    const tsv = buildSelectionTSV(table);
    if (tsv === null) return;

    // 1. Best effort async clipboard (works on https / localhost)
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(tsv).catch(() => {});
    }
    // 2. Reliable fallback via the hidden proxy (works on file://)
    const ta = ensureExcelProxy(table);
    ta.value = tsv;
    ta.focus({ preventScroll: true });
    ta.select();
    try { document.execCommand('copy'); } catch (e) {}
    if (typeof showToast === 'function') showToast('Copied');
}

// Rectangular paste into the manual grid, with single-value fill and undo support.
function pasteExcelData(table, text) {
    if (!excelSelStart || !text) return;

    pushExcelHistory(table); // snapshot BEFORE mutating so Ctrl+Z can revert

    const rows = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    // Drop a single trailing blank line that Excel/most sources append
    if (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();

    const r1 = Math.min(excelSelStart.r, excelSelEnd.r);
    const r2 = Math.max(excelSelStart.r, excelSelEnd.r);
    const c1 = Math.min(excelSelStart.c, excelSelEnd.c);
    const c2 = Math.max(excelSelStart.c, excelSelEnd.c);

    const isSingle = rows.length === 1 && rows[0].split('\t').length === 1;
    let maxR = r1, maxC = c1;

    if (isSingle) {
        // Fill every selected cell with the one value
        const val = rows[0];
        for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
                const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
                if (cell && !cell.classList.contains('excel-idx')) {
                    cell.innerText = val;
                    updateInternalData(cell);
                }
            }
        }
        maxR = r2; maxC = c2;
    } else {
        rows.forEach((rowStr, rOff) => {
            const cols = rowStr.split('\t');
            cols.forEach((val, cOff) => {
                const tr = r1 + rOff, tc = c1 + cOff;
                const cell = table.querySelector(`td[data-r="${tr}"][data-c="${tc}"]`);
                if (cell && !cell.classList.contains('excel-idx')) {
                    cell.innerText = val;
                    updateInternalData(cell);
                    if (tr > maxR) maxR = tr;
                    if (tc > maxC) maxC = tc;
                }
            });
        });
    }

    // Expand selection to cover the pasted block
    excelSelStart = { r: r1, c: c1 };
    excelSelEnd = { r: maxR, c: maxC };
    updateLegacyGlobals(r1, c1, maxR, maxC);
    highlightExcelRange(table);
    if (typeof showToast === 'function') showToast('Pasted');
}

// ============================================================
// EXCEL EDIT-MODE ENGINE: proxy, keyboard, history, editing
// ============================================================

// Create (once) the hidden textarea that captures native copy/cut/paste events.
// This is what makes Ctrl+C / Ctrl+V reliable even on file:// pages.
function ensureExcelProxy(table) {
    if (excelClipboardProxy && document.body.contains(excelClipboardProxy)) return excelClipboardProxy;

    const ta = document.createElement('textarea');
    ta.id = 'excelClipboardProxy';
    ta.setAttribute('aria-hidden', 'true');
    ta.tabIndex = -1;
    ta.style.cssText = 'position:fixed; left:-9999px; top:0; width:1px; height:1px; opacity:0; padding:0; border:0; z-index:-1;';
    document.body.appendChild(ta);
    excelClipboardProxy = ta;

    const getTable = () => document.getElementById('manualRawTable');

    // COPY -> serve a fresh rectangular TSV of the current selection
    ta.addEventListener('copy', function (e) {
        const tsv = buildSelectionTSV(getTable());
        if (tsv === null) return;
        e.preventDefault();
        (e.clipboardData || window.clipboardData).setData('text/plain', tsv);
    });

    // CUT -> copy then clear the block
    ta.addEventListener('cut', function (e) {
        const t = getTable();
        const tsv = buildSelectionTSV(t);
        if (tsv === null) return;
        e.preventDefault();
        (e.clipboardData || window.clipboardData).setData('text/plain', tsv);
        clearSelectedCells(t);
    });

    // PASTE -> rectangular grid paste into the selected block
    ta.addEventListener('paste', function (e) {
        const t = getTable();
        if (!excelSelStart) return;
        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text');
        pasteExcelData(t, text);
        setTimeout(() => syncExcelClipboard(t), 0);
    });

    // KEYS -> undo/redo, delete, arrows, type-to-edit
    ta.addEventListener('keydown', function (e) {
        handleExcelProxyKey(e, getTable());
    });

    return ta;
}

// Load the current selection into the proxy and give it focus, so the browser's
// own Ctrl+C / Ctrl+V / Ctrl+Z land on our handlers. Only active in Edit Mode.
function syncExcelClipboard(table) {
    if (!isEditMode || !table) return;
    // Don't steal focus while the user is typing directly into a cell
    if (document.activeElement && document.activeElement.isContentEditable) return;
    const ta = ensureExcelProxy(table);
    ta.value = buildSelectionTSV(table) || '';
    ta.focus({ preventScroll: true });
    ta.select();
}

// Shared keyboard handler (used by both the proxy and the table in Edit Mode)
function handleExcelProxyKey(e, table) {
    if (!table || !excelSelStart) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const onProxy = document.activeElement === excelClipboardProxy;

    // UNDO / REDO
    if (ctrl && (e.key === 'z' || e.key === 'Z')) {
        e.preventDefault();
        if (e.shiftKey) redoExcel(table); else undoExcel(table);
        return;
    }
    if (ctrl && (e.key === 'y' || e.key === 'Y')) {
        e.preventDefault();
        redoExcel(table);
        return;
    }

    // COPY / CUT / PASTE
    if (ctrl && (e.key === 'c' || e.key === 'C')) {
        if (onProxy) return;            // let the native copy listener fire
        e.preventDefault(); copyExcelData(table); return;
    }
    if (ctrl && (e.key === 'x' || e.key === 'X')) {
        if (onProxy) return;            // let the native cut listener fire
        e.preventDefault(); copyExcelData(table); clearSelectedCells(table); return;
    }
    if (ctrl && (e.key === 'v' || e.key === 'V')) {
        if (onProxy) return;            // let the native paste listener fire
        e.preventDefault();
        if (navigator.clipboard && navigator.clipboard.readText) {
            navigator.clipboard.readText().then(txt => pasteExcelData(table, txt)).catch(() => {});
        }
        return;
    }

    // SELECT ALL
    if (ctrl && (e.key === 'a' || e.key === 'A')) {
        e.preventDefault(); selectAllExcel(table); return;
    }

    // DELETE / BACKSPACE -> clear block
    if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault(); clearSelectedCells(table); return;
    }

    // ARROWS (Shift extends the selection, like Excel)
    if (e.key.startsWith('Arrow')) {
        e.preventDefault();
        moveExcelSelectionEx(e.key, table, e.shiftKey);
        syncExcelClipboard(table);
        return;
    }

    // ENTER -> move down one cell
    if (e.key === 'Enter') {
        e.preventDefault();
        moveExcelSelectionEx('ArrowDown', table, false);
        syncExcelClipboard(table);
        return;
    }

    // F2 -> edit active cell
    if (e.key === 'F2') {
        e.preventDefault();
        const cell = table.querySelector(`td[data-r="${excelSelEnd.r}"][data-c="${excelSelEnd.c}"]`);
        if (cell) beginCellEdit(cell, null);
        return;
    }

    // Any printable character -> start editing the active cell, overwriting it
    if (!ctrl && !e.altKey && e.key.length === 1) {
        const cell = table.querySelector(`td[data-r="${excelSelEnd.r}"][data-c="${excelSelEnd.c}"]`);
        if (cell && !cell.classList.contains('excel-idx')) {
            e.preventDefault();
            beginCellEdit(cell, e.key);
        }
    }
}

// Arrow-key movement with optional Shift-extend
function moveExcelSelectionEx(key, table, extend) {
    let r = excelSelEnd.r, c = excelSelEnd.c;
    if (key === 'ArrowUp') r--;
    if (key === 'ArrowDown') r++;
    if (key === 'ArrowLeft') c--;
    if (key === 'ArrowRight') c++;
    if (r < 0) r = 0;
    if (c < 0) c = 0;

    // Stay within the rendered grid
    if (!table.querySelector(`td[data-r="${r}"][data-c="${c}"]`)) return;

    excelSelEnd = { r, c };
    if (!extend) excelSelStart = { r, c };
    updateLegacyGlobals(excelSelStart.r, excelSelStart.c, excelSelEnd.r, excelSelEnd.c);
    highlightExcelRange(table);

    const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
    if (cell) cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

// Select the whole used range of the grid
function selectAllExcel(table) {
    let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0, found = false;
    table.querySelectorAll('td[data-r]').forEach(cell => {
        if (cell.classList.contains('excel-idx')) return;
        const r = parseInt(cell.getAttribute('data-r'));
        const c = parseInt(cell.getAttribute('data-c'));
        if (isNaN(r) || isNaN(c)) return;
        found = true;
        minR = Math.min(minR, r); minC = Math.min(minC, c);
        maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
    });
    if (!found) return;
    excelSelStart = { r: minR, c: minC };
    excelSelEnd = { r: maxR, c: maxC };
    updateLegacyGlobals(minR, minC, maxR, maxC);
    highlightExcelRange(table);
    syncExcelClipboard(table);
}

// Clear every cell in the current selection
function clearSelectedCells(table) {
    if (!excelSelStart || !excelSelEnd) return;
    pushExcelHistory(table);
    const r1 = Math.min(excelSelStart.r, excelSelEnd.r), r2 = Math.max(excelSelStart.r, excelSelEnd.r);
    const c1 = Math.min(excelSelStart.c, excelSelEnd.c), c2 = Math.max(excelSelStart.c, excelSelEnd.c);
    for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
            const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if (cell && !cell.classList.contains('excel-idx')) {
                cell.innerText = '';
                updateInternalData(cell);
            }
        }
    }
    syncExcelClipboard(table);
}

// Open one cell for typing (double-click, F2, or start typing)
function beginCellEdit(cell, initialChar) {
    const table = cell.closest('table');
    pushExcelHistory(table);

    cell.contentEditable = true;
    cell.classList.add('excel-editing');
    if (initialChar !== null) cell.innerText = initialChar;
    cell.focus();

    // Put the caret at the end of the text
    try {
        const range = document.createRange();
        range.selectNodeContents(cell);
        range.collapse(false);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
    } catch (e) {}

    const finish = function () {
        cell.contentEditable = false;
        cell.classList.remove('excel-editing');
        updateInternalData(cell);
        cell.removeEventListener('blur', finish);
        cell.onkeydown = null;
        // Return focus to the proxy so keyboard shortcuts keep working
        if (table) setTimeout(() => syncExcelClipboard(table), 0);
    };
    cell.addEventListener('blur', finish);

    cell.onkeydown = function (ev) {
        if (ev.key === 'Enter' && !ev.shiftKey) { ev.preventDefault(); cell.blur(); }
        else if (ev.key === 'Escape') { ev.preventDefault(); cell.blur(); }
        else { ev.stopPropagation(); } // keep proxy shortcuts from firing mid-edit
    };
}

// Undo/redo for the manual grid, implemented as whole-grid text snapshots.
function snapshotExcelTable(table) {
    const snap = {};
    table.querySelectorAll('td[data-r]').forEach(cell => {
        if (cell.classList.contains('excel-idx')) return;
        snap[cell.getAttribute('data-r') + '-' + cell.getAttribute('data-c')] = cell.innerText;
    });
    return snap;
}

function applyExcelSnapshot(table, snap) {
    Object.keys(snap).forEach(key => {
        const dash = key.indexOf('-');
        const r = key.slice(0, dash), c = key.slice(dash + 1);
        const cell = table.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
        if (cell && cell.innerText !== snap[key]) {
            cell.innerText = snap[key];
            commitCellValue(cell); // write to workbook without the green flash
        }
    });
}

function pushExcelHistory(table) {
    if (!table) return;
    excelHistory.push(snapshotExcelTable(table));
    if (excelHistory.length > 50) excelHistory.shift(); // cap memory
    excelRedo = []; // a new edit invalidates the redo stack
}

function undoExcel(table) {
    if (excelHistory.length === 0) { if (typeof showToast === 'function') showToast('Nothing to undo'); return; }
    excelRedo.push(snapshotExcelTable(table));
    applyExcelSnapshot(table, excelHistory.pop());
    syncExcelClipboard(table);
    if (typeof showToast === 'function') showToast('Undo');
}

function redoExcel(table) {
    if (excelRedo.length === 0) { if (typeof showToast === 'function') showToast('Nothing to redo'); return; }
    excelHistory.push(snapshotExcelTable(table));
    applyExcelSnapshot(table, excelRedo.pop());
    syncExcelClipboard(table);
    if (typeof showToast === 'function') showToast('Redo');
}

// Moves the single-cell selection with the arrow keys (non-edit mode).
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
    if (cell) cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

//STEP-2 EDIITING HELPER FUNCTIONS
function handleHeaderEdit(side, colIdx, newName) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    if (data && data.headers) {
        data.headers[colIdx] = newName.trim();
    }
}

function handleCellEdit(side, rIdx, cIdx, newVal) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    if (data && data.body && data.body[rIdx]) {
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

// ==========================================
// STEP 2 SPREADSHEET: select / edit / navigate cells
// ==========================================

// Selects a single cell in a Step 2 table.
function selectExcelCell(side, r, c) {
    if (excelState.editing) return;
    excelState = { side: side, mode: 'cell', r: r, c: c, editing: false };
    renderPreviewTables();
}

// Selects (or toggles off) an entire column in a Step 2 table.
function selectExcelCol(side, c) {
    if (excelState.side === side && excelState.mode === 'col' && excelState.c === c) {
        excelState = { side: null, mode: null, r: -1, c: -1, editing: false };
    } else {
        excelState = { side: side, mode: 'col', r: -1, c: c, editing: false };
    }
    renderPreviewTables();
}

// Opens a Step 2 cell for typing (double-click or start typing).
function editExcelCell(side, r, c, cellEl, initialValue = null) {
    excelState.editing = true;
    cellEl.contentEditable = true;
    cellEl.classList.add('excel-editing');

    // If a keystroke started the edit, overwrite the cell with that character.
    if (initialValue !== null) {
        cellEl.innerText = initialValue;
        // Put the caret at the end.
        const range = document.createRange();
        const sel = window.getSelection();
        range.selectNodeContents(cellEl);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
    }

    cellEl.focus();

    // Commit on blur, and on Enter (which blurs).
    cellEl.onblur = function () {
        saveExcelEdit(side, r, c, this.innerText);
    };
    cellEl.onkeydown = function (e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            this.blur();
        }
    };
}

// Writes an edited Step 2 cell back to the model and re-renders.
function saveExcelEdit(side, r, c, val) {
    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;
    if (data && data.body[r]) {
        data.body[r][c] = val.trim();
    }
    excelState.editing = false;
    renderPreviewTables();
}

// Keyboard handler for the Step 2 grid: navigation, delete, and type-to-edit.
function handleExcelKey(e, side) {
    // While editing a cell, only intercept Enter/Escape/Tab.
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
            excelState.c = Math.min(excelState.c + 1, projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'].headers.length - 1);
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

    // Shortcuts (when not editing):

    // Delete/Backspace clears the cell, or the whole column if one is selected.
    if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();

        if (excelState.mode === 'col') {
            if (confirm("Clear entire column?")) {
                for (let i = 0; i <= maxR; i++) data.body[i][c] = "";
                showToast("Column Cleared");
            }
        }
        else if (excelState.mode === 'cell') {
            data.body[r][c] = "";
        }
        renderPreviewTables();
        return;
    }

    // Arrow keys / Tab move the active cell.
    if (e.key === 'ArrowUp') { r--; e.preventDefault(); }
    else if (e.key === 'ArrowDown') { r++; e.preventDefault(); }
    else if (e.key === 'ArrowLeft') { c--; e.preventDefault(); }
    else if (e.key === 'ArrowRight') { c++; e.preventDefault(); }
    else if (e.key === 'Tab') { c++; e.preventDefault(); }

    // Enter opens the active cell for editing.
    else if (e.key === 'Enter') {
        e.preventDefault();
        const cell = document.querySelector(`#gridContainer-${side} .excel-focus`);
        if (cell) editExcelCell(side, r, c, cell);
        return;
    }

    // A printable key starts editing and overwrites the cell.
    else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        const cell = document.querySelector(`#gridContainer-${side} .excel-focus`);
        if (cell) {
            editExcelCell(side, r, c, cell, e.key);
            e.preventDefault();
            return;
        }
    }

    // Clamp the new position to the grid.
    if (r < 0) r = 0; if (r > maxR) r = maxR;
    if (c < 0) c = 0; if (c > maxC) c = maxC;

    // Re-render only if the selection moved.
    if (r !== excelState.r || c !== excelState.c) {
        excelState.r = r;
        excelState.c = c;
        renderPreviewTables();
    }
}

// Paste into the Step 2 grid: fills a selected column as a pattern, or pastes a
// block starting at the active cell.
function handleExcelPaste(e, side) {
    e.preventDefault();
    e.stopPropagation();

    // Exit edit mode first so paste lands on the grid, not inside one cell.
    if (excelState.editing) {
        document.activeElement.blur();
        excelState.editing = false;
    }

    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    const rows = text.split(/\r\n|\n|\r/).filter(r => r.trim());
    if (rows.length === 0) return;

    const p = projects[activeProjectIdx];
    const data = side === 'A' ? p.dataA : p.dataB;

    // A whole column is selected -> repeat the pasted values down the column.
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

                if (data.body[targetR] && targetC < data.headers.length) {
                    data.body[targetR][targetC] = val.trim();
                }
            });
        });
        showToast(`Pasted ${rows.length} rows.`);
    }

    renderPreviewTables();
}

// Makes a Step 2 column header editable in place; saves the new name on blur/Enter.
function editHeader(side, colIdx, th) {
    th.contentEditable = true;
    th.focus();
    th.onblur = function () {
        const p = projects[activeProjectIdx];
        const data = side === 'A' ? p.dataA : p.dataB;
        data.headers[colIdx] = this.innerText.trim();
        th.contentEditable = false;
    };
    th.onkeydown = (e) => { if (e.key === 'Enter') th.blur(); };
}

// Collapses/expands the left sidebar and swaps the toggle icon.
function toggleSidebar() {
    const sidebar = document.querySelector('.sidebar');
    const icon = document.querySelector('#sidebarToggle i');

    sidebar.classList.toggle('collapsed');

    if (sidebar.classList.contains('collapsed')) {
        icon.className = 'fas fa-chevron-right';
    } else {
        icon.className = 'fas fa-bars';
    }
}

// ==========================================
// GRID ENGINE — the full Step 2 spreadsheet
// Selection, auto-scroll, copy/paste, undo/redo and the fill handle live here.
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

    init: function (side) {
        const container = document.getElementById(`gridContainer-${side}`);
        if (!container) return;

        const table = container.querySelector('.clean-table');
        const scrollable = container.closest('.scroll-wrap');
        const fillHandle = document.getElementById(`fillHandle${side}`);

        // Allow native text selection only inside a cell that's being edited.
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

                // Keep the horizontal scroll position steady while selecting.
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

    handleGlobalMouseMove: function (e) {
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

    handleGlobalMouseUp: function (e) {
        if (this.isDraggingFill && this.activeSide) this.executeFill(this.activeSide);
        this.isDragging = false;
        this.isDraggingFill = false;
        this.stopAutoScroll();
    },

    startAutoScroll: function (scrollable, table, fillHandle, side) {
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

    stopAutoScroll: function () {
        if (this.autoScrollTimer) clearInterval(this.autoScrollTimer);
        this.autoScrollTimer = null;
    },

    updateSelection: function (start, end, table, fillHandle) {
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

    handleKeyDown: function (e, table, side, fillHandle) {
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

    handleCopy: function (e, table) {
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
        const endR = Math.max(this.selStart.r, this.selEnd.r);
        const startC = Math.min(this.selStart.c, this.selEnd.c);
        const endC = Math.max(this.selStart.c, this.selEnd.c);
        
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];
        let rowsUpdated = 0;

        // If one value was copied but many cells are selected, fill them all.
        const isSingleValue = pasteRows.length === 1 && pasteRows[0].split('\t').length === 1;
        const isMultiCellSelection = (startR !== endR) || (startC !== endC);

        if (isSingleValue && isMultiCellSelection) {
            const fillValue = pasteRows[0].trim();
            for (let r = startR; r <= endR; r++) {
                for (let c = startC; c <= endC; c++) {
                    if (data.body[r] && c < data.headers.length) {
                        data.body[r][c] = fillValue;
                        rowsUpdated++;
                    }
                }
            }
        } else {
            // Otherwise paste the grid starting at the top-left of the selection.
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
        }

        renderPreviewTables();
        showToast(`Pasted to ${rowsUpdated} cells!`);
    },

    clearSelection: function (side) {
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

    drawFillPreview: function (table) {
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

    executeFill: function (side) {
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

    saveState: function (side) {
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

    undo: function (side) {
        if (this.history[side].undo.length === 0) return showToast("Nothing to undo");
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];

        const currentSnapshot = data.body.map(r => { const nr = [...r]; Object.defineProperty(nr, '_originalIdx', { value: r._originalIdx, writable: true, enumerable: false }); return nr; });
        this.history[side].redo.push(currentSnapshot);

        data.body = this.history[side].undo.pop();
        if (this.history[side].undo.length === 0) document.getElementById(`undoBadge${side}`).style.display = 'none';
        renderPreviewTables();
        showToast("Undid last action");
    },

    redo: function (side) {
        if (this.history[side].redo.length === 0) return showToast("Nothing to redo");
        const data = projects[activeProjectIdx][side === 'A' ? 'dataA' : 'dataB'];

        const currentSnapshot = data.body.map(r => { const nr = [...r]; Object.defineProperty(nr, '_originalIdx', { value: r._originalIdx, writable: true, enumerable: false }); return nr; });
        this.history[side].undo.push(currentSnapshot);

        data.body = this.history[side].redo.pop();
        document.getElementById(`undoBadge${side}`).style.display = 'inline';
        renderPreviewTables();
        showToast("Redid action");
    },

    startEditing: function (cell) {
        this.editingCell = cell;
        cell.contentEditable = true;
        cell.classList.add('excel-editing');
        cell.focus();
        cell.setAttribute('data-original', cell.innerText);
        cell.onblur = () => { if (this.editingCell === cell) this.stopEditing(true, this.activeSide); };
    },

    stopEditing: function (saveValue, side) {
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
        // Refocus without scrolling, so finishing an edit doesn't jump the view.
        if (container) container.focus({ preventScroll: true });
    }
};

// ==========================================
// FILE QUEUE & REDIRECT (Table 1 direct / Table 2 queue)
// ==========================================

let table2Queue = [];

// Table 1 upload entry point: tags the batch so identical re-uploads still group apart.
function handleTable1Direct(input) {
    if (!input.files || input.files.length === 0) return;

    input.uploadSessionId = Date.now();

    handleBulkUpload(input);
}

// Table 2 upload: real CPQ files are queued; "default"/empty files are redirected
// to the Table 1 (Form) importer instead.
function handleTable2Queue(input) {
    if (!input.files || input.files.length === 0) return;

    const incomingFiles = Array.from(input.files);
    let redirectedFiles = [];

    incomingFiles.forEach(file => {
        const isDefault = file.name.toLowerCase().includes('default') || file.size === 0;

        if (isDefault) {
            redirectedFiles.push(file);
        } else {
            table2Queue.push({ fileData: file, status: 'Pending' });
        }
    });

    if (redirectedFiles.length > 0) {
        // Pass a plain {files} object instead of a DataTransfer.
        handleBulkUpload({ files: redirectedFiles, value: "", uploadSessionId: Date.now() });

        showToast(`${redirectedFiles.length} default files redirected to Table 1!`);
    }

    renderTable2QueueUI();
    input.value = "";
}

// Uploads a single queued CPQ file into the next Set.
function processQueueFile(index) {
    const queueItem = table2Queue[index];
    if (queueItem.status === 'Uploaded ✅') return;

    handleCPQUpload({ files: [queueItem.fileData], value: "" });

    queueItem.status = 'Uploaded ✅';
    renderTable2QueueUI();
}

// Uploads all pending queued CPQ files at once.
function processAllQueue() {
    const pendingItems = table2Queue.filter(item => item.status === 'Pending');
    if (pendingItems.length === 0) return;

    const filesToUpload = pendingItems.map(item => {
        item.status = 'Uploaded ✅';
        return item.fileData;
    });

    handleCPQUpload({ files: filesToUpload, value: "" });

    renderTable2QueueUI();
    showToast(`Processed ${pendingItems.length} queued files!`);
}

// Renders the Table 2 queue list UI.
function renderTable2QueueUI() {
    const tbody = document.getElementById('table2QueueList');
    const processAllBtn = document.getElementById('btnProcessAllQueue');
    if (!tbody) return;
    
    if (table2Queue.length === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:#94a3b8; padding: 5px;">Queue is empty.</td></tr>';
        if (processAllBtn) processAllBtn.style.display = 'none';
        return;
    }

    const hasPending = table2Queue.some(item => item.status === 'Pending');
    if (processAllBtn) processAllBtn.style.display = hasPending ? 'inline-block' : 'none';

    tbody.innerHTML = table2Queue.map((item, index) => {
        const isDone = item.status.includes('✅');
        return `
        <tr>
            <td style="color:#334155; padding: 5px; max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.fileData.name}">
                ${item.fileData.name}
            </td>
            <td style="color:${isDone ? '#10b981' : '#f59e0b'}; font-weight:bold; padding: 5px;">${item.status}</td>
            <td style="text-align:right; padding: 5px;">
                ${!isDone 
                    ? `<button onclick="processQueueFile(${index})" style="background:#2563eb; color:white; border:none; padding:3px 8px; border-radius:3px; cursor:pointer; font-size:10px;">Process</button>` 
                    : `<span style="color:#94a3b8; font-size:10px;">Done</span>`
                }
            </td>
        </tr>`;
    }).join('');
}

// ==========================================
// LINKED FOLDER (File System Access API)
// Auto-fetch the newest Excel files from a folder the user links once.
// ==========================================

let linkedFolderHandle = null;

// IndexedDB is used to remember the linked folder handle across page refreshes.
const DB_NAME = "FolderMemoryDB";
function getDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = e => e.target.result.createObjectStore('handles');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function saveFolderHandle(handle) {
    const db = await getDB();
    db.transaction('handles', 'readwrite').objectStore('handles').put(handle, 'linkedFolder');
}

async function loadFolderHandle() {
    const db = await getDB();
    return new Promise((resolve) => {
        const req = db.transaction('handles', 'readonly').objectStore('handles').get('linkedFolder');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
    });
}

// On load, restore the linked-folder label if a handle was saved.
window.addEventListener('DOMContentLoaded', async () => {
    linkedFolderHandle = await loadFolderHandle();
    if (linkedFolderHandle) {
        const btnText = document.getElementById('lblLinkedFolder');
        if (btnText) btnText.innerText = "Linked: " + linkedFolderHandle.name;
    }
});

// Ensures we still have read permission (may prompt after a refresh).
async function verifyPermission(fileHandle) {
    if ((await fileHandle.queryPermission({ mode: 'read' })) === 'granted') return true;
    if ((await fileHandle.requestPermission({ mode: 'read' })) === 'granted') return true;
    return false;
}

// Prompts the user to pick a folder and remembers it.
async function linkLocalFolder() {
    try {
        linkedFolderHandle = await window.showDirectoryPicker({ mode: 'read' });
        await saveFolderHandle(linkedFolderHandle);

        const btnText = document.getElementById('lblLinkedFolder');
        if (btnText) btnText.innerText = "Linked: " + linkedFolderHandle.name;
        showToast("Successfully linked to " + linkedFolderHandle.name);
    } catch (err) {
        console.warn("Folder linking cancelled:", err);
    }
}

// Returns the `limit` most recently modified Excel/CSV files from the linked folder.
async function getNewestExcelFiles(limit) {
    if (!linkedFolderHandle) {
        showModal("Folder Not Linked", "Please click 'Link Folder' first.", "warning");
        return [];
    }

    const hasPermission = await verifyPermission(linkedFolderHandle);
    if (!hasPermission) {
        showModal("Permission Denied", "Please click 'Link Folder' again to restore access.", "error");
        return [];
    }

    let excelFiles = [];
    try {
        for await (const entry of linkedFolderHandle.values()) {
            if (entry.kind === 'file' && entry.name.match(/\.(xlsx|xls|csv)$/i)) {
                const fileData = await entry.getFile();
                if (fileData.size > 0 && !entry.name.startsWith('~$')) {
                    excelFiles.push(fileData);
                }
            }
        }
        excelFiles.sort((a, b) => b.lastModified - a.lastModified);
        return excelFiles.slice(0, limit);
    } catch (err) {
        console.error("Error reading directory:", err);
        return [];
    }
}

// Auto-loads the newest Form files from the linked folder into Table 1.
async function autoLoadTable1() {
    const qtyInput = document.getElementById('autoFileQtyA');
    const limit = qtyInput ? parseInt(qtyInput.value) : 1;

    const files = await getNewestExcelFiles(limit);
    if (files.length === 0) {
        showModal("No Files Found", "Could not find any recent Excel files.", "error");
        return;
    }

    showToast(`Loading ${files.length} latest file(s)...`);

    handleBulkUpload({ files: files, value: "" });
}

// Auto-loads the newest CPQ files from the linked folder into the Table 2 queue.
async function autoLoadTable2() {
    const qtyInput = document.getElementById('autoFileQty');
    const limit = qtyInput ? parseInt(qtyInput.value) : 10;
    
    const files = await getNewestExcelFiles(limit);
    if (files.length === 0) return;

    showToast(`Grabbing the ${files.length} most recent files...`);
    
    handleTable2Queue({ files: files, value: "" });
    
    setTimeout(() => {
        if (typeof processAllQueue === 'function') processAllQueue();
    }, 100);
}

// ==========================================
// MASTER AUTO-BATCH
// Pulls files from the linked folder, splits them into Forms vs CPQs (by filename),
// creates the Sets, and pairs each CPQ to its Set.
// ==========================================

// Classifies linked-folder files into Forms and CPQs and kicks off the batch.
// CPQ filenames match a "1234_5"-style pattern.
async function runMasterBatch(mode = 'auto') {
    let formFiles = [];
    let cpqFiles = [];
    const cpqRegex = /\d{4,}_\d/;

    if (mode === 'auto') {
        const totalQty = parseInt(document.getElementById('batchTotalQty').value) || 2;
        // Grab extra files so hidden Windows temp files don't crowd out real ones.
        const pool = await getNewestExcelFiles(totalQty + 20);
        if (pool.length === 0) return showModal("No Files", "Please link your local folder.", "error");

        let validFiles = [];
        for (let f of pool) {
            if (!f.name.startsWith('~$')) validFiles.push(f);
        }
        validFiles = validFiles.slice(0, totalQty);

        for (let file of validFiles) {
            if (cpqRegex.test(file.name)) cpqFiles.push(file);
            else formFiles.push(file);
        }
    } else {
        const numForms = parseInt(document.getElementById('batchFormQty').value) || 1;
        const numCPQs = parseInt(document.getElementById('batchCpqQty').value) || 1;
        const pool = await getNewestExcelFiles(Math.max(50, numForms + numCPQs + 20));

        if (pool.length === 0) return showModal("No Files", "Please link your local folder.", "error");

        for (let file of pool) {
            if (file.name.startsWith('~$')) continue;
            if (cpqRegex.test(file.name)) {
                if (cpqFiles.length < numCPQs) cpqFiles.push(file);
            } else {
                if (formFiles.length < numForms) formFiles.push(file);
            }
        }
    }

    if (formFiles.length === 0) {
        showModal("Missing Forms", "Could not find any Order Form files in that batch.", "error");
        return;
    }

    formFiles.sort((a, b) => a.lastModified - b.lastModified);
    cpqFiles.sort((a, b) => a.lastModified - b.lastModified);

    showToast(`Batching ${formFiles.length} Forms and ${cpqFiles.length} CPQs...`);
    window.pendingBatchCPQs = cpqFiles;

    handleBulkUpload({ files: formFiles, value: "" });
}

// Reads a File into a Uint8Array (promise-wrapped FileReader).
function readExcelAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(new Uint8Array(e.target.result));
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Injects each CPQ file's data into the matching newly-created Set (by order).
async function processBatchCPQs(cpqFiles, newSetCount) {
    const startIdx = projects.length - newSetCount;
    showToast(`Injecting CPQ files into ${newSetCount} tabs...`);

    for (let i = 0; i < newSetCount; i++) {
        if (i >= cpqFiles.length) break;
        const pIdx = startIdx + i;
        const p = projects[pIdx];
        const file = cpqFiles[i];

        try {
            const data = await readExcelAsArrayBuffer(file);
            const workbook = XLSX.read(data, { type: 'array' });

            let extracted = "";

            // Prefer sheet 2's SKU table (dropping its first column).
            if (workbook.SheetNames.length > 1) {
                const sheet2 = workbook.Sheets[workbook.SheetNames[1]];
                const rows = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
                let start = -1;
                for (let r = 0; r < rows.length; r++) {
                    if ((rows[r] || []).join(" ").toLowerCase().includes("sku information")) {
                        start = r + 1;
                        break;
                    }
                }
                if (start !== -1) {
                    extracted = rows.slice(start)
                        .map(r => r.slice(1).map(c => (c == null) ? "" : c).join("\t"))
                        .join("\n");
                }
            }

            // Otherwise fall back to all of sheet 1.
            if (!extracted) {
                const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
                extracted = rows.map(r => r.map(c => (c == null) ? "" : c).join("\t")).join("\n");
            }

            if (extracted) {
                p.rawB = extracted;
                p.fileNameB = file.name;
                p.fileBName = file.name; 
                p.status = 'ready'; 
            }
        } catch (err) {
            console.error(`Error injecting CPQ ${file.name}:`, err);
        }
    }
    
    renderTopBar();
    switchProject(startIdx, true); 
    showModal("Master Batch Complete", `✅ Successfully built and paired <strong>${newSetCount}</strong> sets!`, "success");
}

// ==========================================
// STEP 2 DATA TOOLS (merge columns, bulk text edit)
// ==========================================

// Merges the highlighted columns into the leftmost one (e.g. to reassemble a
// barcode split across cells) and removes the now-empty columns.
function stitchColumns(side) {
    let start = typeof GridEngine !== 'undefined' ? GridEngine.selStart : excelSelStart;
    let end = typeof GridEngine !== 'undefined' ? GridEngine.selEnd : excelSelEnd;

    if (!start || !end) return showModal("No Selection", "Please click and drag to highlight the columns you want to merge.", "warning");

    let cMin = Math.min(start.c, end.c);
    let cMax = Math.max(start.c, end.c);

    if (cMin === cMax) return showModal("Invalid Selection", "Please highlight at least TWO columns to stitch them together.", "warning");

    const p = projects[activeProjectIdx];
    const data = (side === 'A') ? p.dataA : p.dataB;

    // Snapshot for undo.
    if (typeof GridEngine !== 'undefined' && GridEngine.saveState) GridEngine.saveState(side);

    // Concatenate the columns in each active row, then drop the extra columns.
    data.body.forEach(row => {
        let stitchedValue = "";
        for (let c = cMin; c <= cMax; c++) {
            stitchedValue += String(row[c] || "").trim();
        }
        row[cMin] = stitchedValue;
        row.splice(cMin + 1, cMax - cMin);
    });

    // Do the same for hidden rows so they stay column-aligned.
    if (data.hiddenRows) {
        data.hiddenRows.forEach(item => {
            let row = item.data;
            let stitchedValue = "";
            for (let c = cMin; c <= cMax; c++) {
                stitchedValue += String(row[c] || "").trim();
            }
            row[cMin] = stitchedValue;
            row.splice(cMin + 1, cMax - cMin);
        });
    }

    // Rename the merged column and drop the now-empty header slots.
    let newHeaderName = data.headers[cMin] || "Merged Data";
    if (newHeaderName.toLowerCase().startsWith("column ")) newHeaderName = "Merged Barcode";
    data.headers[cMin] = newHeaderName;
    data.headers.splice(cMin + 1, cMax - cMin);

    if (typeof GridEngine !== 'undefined') { GridEngine.selStart = null; GridEngine.selEnd = null; }
    excelSelStart = null; excelSelEnd = null;

    renderPreviewTables();
    showToast(`Stitched ${cMax - cMin + 1} columns into one!`);
}

// Adds a prefix to, or removes a substring from, every highlighted cell.
function bulkTextEdit(side, action) {
    let start = typeof GridEngine !== 'undefined' ? GridEngine.selStart : excelSelStart;
    let end = typeof GridEngine !== 'undefined' ? GridEngine.selEnd : excelSelEnd;

    if (!start || !end) {
        return showModal("No Selection", "Please click and drag to highlight the specific cells or column you want to edit.", "warning");
    }

    // Ask what to add/remove.
    let promptMsg = action === 'add'
        ? "What text do you want to ADD to the front of these cells? (e.g. '0' or 'USD ')"
        : "What exact text do you want to REMOVE from these cells? (e.g. '0' or 'Pcs')";

    let userInput = prompt(promptMsg);
    if (!userInput) return;

    let rMin = Math.min(start.r, end.r);
    let rMax = Math.max(start.r, end.r);
    let cMin = Math.min(start.c, end.c);
    let cMax = Math.max(start.c, end.c);

    const p = projects[activeProjectIdx];
    const data = (side === 'A') ? p.dataA : p.dataB;

    // Snapshot for undo.
    if (typeof GridEngine !== 'undefined' && GridEngine.saveState) GridEngine.saveState(side);

    let changedCount = 0;

    for (let r = rMin; r <= rMax; r++) {
        for (let c = cMin; c <= cMax; c++) {
            if (data.body[r] && data.body[r][c] !== undefined) {
                let val = String(data.body[r][c]).trim();

                if (val !== "") {
                    if (action === 'add') {
                        data.body[r][c] = userInput + val;
                        changedCount++;
                    } else if (action === 'remove') {
                        if (val.includes(userInput)) {
                            // split/join removes every occurrence.
                            data.body[r][c] = val.split(userInput).join('').trim();
                            changedCount++;
                        }
                    }
                }
            }
        }
    }

    renderPreviewTables();
    showToast(`Updated ${changedCount} cells successfully!`);
}

// ==========================================
// LIVE TRANSLATION (Chinese -> English)
// ==========================================

// Translates Chinese text in a Set's data/headers to English, backing up the
// originals so the change can be undone.
async function universalTranslate(targetSide = null) {
    const p = projects[activeProjectIdx];
    if (!p || !p.dataA || !p.dataB) return;

    const hasChinese = /[\u4e00-\u9fff]/;
    let translatedCount = 0;

    showToast("🌍 Translating... Please wait.");

    // Translates any Chinese entries in a header array in place (backing them up first).
    async function translateArray(headerArray, side) {
        p[`backupHeaders${side}`] = [...headerArray];

        for (let i = 0; i < headerArray.length; i++) {
            let text = String(headerArray[i] || "").trim();

            if (hasChinese.test(text)) {
                try {
                    let response = await fetch(`https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=zh-CN|en`);
                    let json = await response.json();
                    
                    if (json && json.responseData && json.responseData.translatedText) {
                        let englishResult = json.responseData.translatedText.trim();
                        if (englishResult && englishResult !== text) {
                            headerArray[i] = englishResult;
                            translatedCount++;
                        }
                    }
                } catch (err) {
                    console.error(`Translation failed for: ${text}`, err);
                }
            }
        }
    }

    // Translate the requested side (or both if none specified).
    if (targetSide === 'A' || !targetSide) await translateArray(p.dataA.headers, 'A');
    if (targetSide === 'B' || !targetSide) await translateArray(p.dataB.headers, 'B');

    // Keep mapping names in sync with the translated CPQ headers.
    if (p.mapping && p.mapping.length > 0) {
        p.mapping.forEach(m => {
            if (p.dataB.headers[m.idxB]) m.name = p.dataB.headers[m.idxB];
        });
    }

    // Repaint and reveal the Undo button when something changed.
    if (translatedCount > 0) {
        showToast(`✅ Successfully translated ${translatedCount} column(s)!`);

        if (targetSide) {
            let undoBtn = document.getElementById(`undoTranslateBtn${targetSide}`);
            if (undoBtn) undoBtn.style.display = 'inline-block';
        }

        if (p.step === 2) renderPreviewTables();
        else if (p.step === 3) renderMappingTable();
        else if (p.step === 4) renderDashboard();
    } else {
        showToast("No Chinese characters found to translate.");
    }
}

// Reverts a translated side back to its backed-up original headers.
function undoTranslation(side) {
    const p = projects[activeProjectIdx];

    if (side === 'A' && p.backupHeadersA) {
        p.dataA.headers = [...p.backupHeadersA];
        p.backupHeadersA = null; // clearing the backup hides the Undo button
    }
    if (side === 'B' && p.backupHeadersB) {
        p.dataB.headers = [...p.backupHeadersB];
        p.backupHeadersB = null;
    }

    showToast("⏪ Translation reversed!");

    if (p.step === 2) renderPreviewTables();
}



// ==========================================
// SMART AUTO-SPLIT (multi-PO / multi-item sheets)
// Splits one sheet holding several PO/Item combinations into separate Sets.
// ==========================================
function performSmartAutoSplit(rawData, cleanTable, fileName, sheetName, workbook) {
    if (!cleanTable || cleanTable.length < 2) return 0;

    // Look above the table for a global list of items (e.g. "Item: A & B").
    let globalItems = [""]; // default: a single empty item
    for (let i = 0; i < Math.min(rawData.length, 30); i++) {
        let rowStr = (rawData[i] || []).join(" ");
        let lowerStr = rowStr.toLowerCase();

        // Ignore "description" so it doesn't get mistaken for the item field.
        if (/\bitem\b|customer\s*item|customer\s*code/i.test(lowerStr) && !lowerStr.includes("desc")) {
            let cleanStr = rowStr.replace(/customer\s*item#?|customer\s*code|item\s*code|item#?|please\s*select|\(|\)|:/gi, "").trim();

            // Multiple items separated by & or , become the global item list.
            if (cleanStr.includes("&") || cleanStr.includes(",")) {
                globalItems = cleanStr.split(/[&,]/).map(s => s.trim()).filter(s => s.length > 0);
                break;
            }
        }
    }

    // Locate the Item / PO / GLID columns in the table header.
    let headers = cleanTable[0] || [];
    let poCol = -1, glidCol = -1, itemCol = -1;

    headers.forEach((h, idx) => {
        if (!h) return;
        let text = String(h).toLowerCase().trim();

        // PO column: "PO", "PO#", "Customer PO", "Vendor PO", "Purchase Order".
        if (/\bpo\b|\bpo\s*[#:]|customer\s*po|vendor\s*po|purchase\s*order/i.test(text)) {
            if (poCol === -1) poCol = idx;
        }
        // GLID column: "Glid", "GLID:", "Glid Number".
        if (/\bglid\b/i.test(text)) {
            if (glidCol === -1) glidCol = idx;
        }
        // Item column (ignoring "description" so it isn't mistaken for the item).
        if (/\bitem\b|customer\s*item|customer\s*code/i.test(text) && !text.includes("desc")) {
            if (itemCol === -1) itemCol = idx;
        }
    });

    // Group the data rows by their Item/PO/GLID combination.
    let groups = {}; // key "item|po|glid" -> { po, glid, item, rows }

    for (let i = 1; i < cleanTable.length; i++) {
        let row = cleanTable[i];
        if (!row || row.every(c => !c || String(c).trim() === "")) continue;

        let rowPO = poCol !== -1 ? String(row[poCol] || "").trim() : "";
        let rowGlid = glidCol !== -1 ? String(row[glidCol] || "").trim() : "";
        let rowItem = itemCol !== -1 ? String(row[itemCol] || "").trim() : "";

        let groupKey = `${rowItem}|${rowPO}|${rowGlid}`;
        if (groupKey === "||") groupKey = "Unassigned"; // rows with no routing info

        if (!groups[groupKey]) {
            groups[groupKey] = { po: rowPO, glid: rowGlid, item: rowItem, rows: [] };
        }
        groups[groupKey].rows.push(row);
    }

    // Create one Set per (global item x group) combination.
    let setsCreated = 0;

    globalItems.forEach(globalItem => {
        Object.keys(groups).forEach(key => {
            let group = groups[key];

            // Compose a descriptive tab name from the item/PO/GLID parts.
            let nameParts = [];
            if (globalItem) nameParts.push(globalItem);
            if (group.item) nameParts.push(group.item);
            if (group.po) nameParts.push(`PO:${group.po}`);
            if (group.glid) nameParts.push(`GLID:${group.glid}`);

            let tabName = nameParts.length > 0 ? nameParts.join(" - ") : "Unassigned";
            let finalName = `${sheetName} (${tabName})`;

            let newTableData = [headers, ...group.rows];

            createSet(finalName, workbook, sheetName, fileName);
            let p = projects[projects.length - 1];
            p.rawA = arrayToTSV(newTableData);
            p.status = 'ready';
            p.uploadSessionId = Date.now(); // group this split batch together
            setsCreated++;
        });
    });

    return setsCreated;
}


// ==========================================
// MANUAL SELECTOR: EDIT MODE
// Turns the raw grid into a mini spreadsheet (drag-select, copy/paste, undo, typing).
// ==========================================

let isEditMode = false;

// Toggles Edit Mode on the manual grid and wires up double-click-to-edit.
function toggleEditMode() {
    isEditMode = !isEditMode;
    const btn = document.getElementById('btnEditToggle');
    const table = document.getElementById('manualRawTable');

    if (!table) return console.error("Table not found!");

    btn.innerHTML = isEditMode ? '<i class="fas fa-save"></i> Edit Mode: ON' : '<i class="fas fa-edit"></i> Edit Mode: OFF';
    btn.style.background = isEditMode ? '#10b981' : '#64748b';

    if (isEditMode) {
        // Force the mode to Manual Drag so users can highlight blocks of cells
        document.querySelector('input[name="manMode"][value="manual"]').checked = true;
        window.manualSelectMode = 'manual';

        // Fresh undo history for this editing session
        excelHistory = [];
        excelRedo = [];

        // Double-click a cell to type inside it (full editing helper: history + focus return)
        table.ondblclick = function(e) {
            const cell = e.target.closest('td');
            if (cell && !cell.classList.contains('excel-idx')) beginCellEdit(cell, null);
        };

        // If something is already selected, arm the clipboard proxy right away
        if (excelSelStart) syncExcelClipboard(table);

        showToast("Edit Mode ON: drag to select, Ctrl+C/V to copy-paste, Ctrl+Z to undo, double-click or type to edit.");
    } else {
        // Leaving edit mode: stop capturing double-clicks and release the proxy focus
        table.ondblclick = null;
        if (excelClipboardProxy) excelClipboardProxy.blur();
    }
}

// Fallback paste handler on the grid itself. The hidden clipboard proxy handles
// paste in normal use; this only fires when a table cell is focused: while typing
// in a cell we allow the native single-cell paste, otherwise we grid-paste.
document.getElementById('manualRawTable').addEventListener('paste', function(e) {
    if (!isEditMode || !excelSelStart || !excelSelEnd) return;
    // A cell is open for typing -> let the browser paste plain text into it
    if (e.target && e.target.isContentEditable) return;

    e.preventDefault();
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!text) return;
    pasteExcelData(this, text);
});

function updateInternalData(cell) {
    commitCellValue(cell);
    // Flash green so the user sees the edit registered
    cell.style.backgroundColor = "#dcfce7";
    setTimeout(() => { cell.style.backgroundColor = ""; }, 500);
}

// Write a cell's text back into the active workbook sheet (no visual flash).
function commitCellValue(cell) {
    let rowIdx = parseInt(cell.getAttribute('data-r'));
    let colIdx = parseInt(cell.getAttribute('data-c'));
    
    if (isNaN(rowIdx) || isNaN(colIdx)) return;

    // Prefer the workbook the manual mapper is actually showing (confirmManualImport
    // reads from manualWorkbook), falling back to the last uploaded one.
    const wb = manualWorkbook || lastUploadedWorkbook;
    if (!wb) return;
    const sheet = wb.Sheets[currentSheetName];
    if (!sheet) return;

    const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
    const v = cell.innerText;
    if (v === '') {
        delete sheet[cellRef];
    } else {
        sheet[cellRef] = { t: 's', v: v };
    }

    // Keep the sheet's declared range in sync so edited/added cells aren't clipped
    try {
        if (v !== '') {
            const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { r: rowIdx, c: colIdx }, e: { r: rowIdx, c: colIdx } };
            if (rowIdx < range.s.r) range.s.r = rowIdx;
            if (colIdx < range.s.c) range.s.c = colIdx;
            if (rowIdx > range.e.r) range.e.r = rowIdx;
            if (colIdx > range.e.c) range.e.c = colIdx;
            sheet['!ref'] = XLSX.utils.encode_range(range);
        }
    } catch (e) {}
}

// Highlights the clicked header row in the manual selector and enables Confirm.
function selectHeaderRow(rowIndex) {
    selectedHeaderRowIndex = rowIndex;
    const rows = document.querySelectorAll('.manual-selection-table tr');
    rows.forEach((r, idx) => {
        r.style.backgroundColor = (idx === rowIndex) ? '#dcfce7' : 'transparent';
        r.style.fontWeight = (idx === rowIndex) ? 'bold' : 'normal';
    });
    document.getElementById('btnManualConfirm').disabled = false;
    document.getElementById('btnManualConfirm').style.background = '#2563eb';
    document.getElementById('btnManualConfirm').style.cursor = 'pointer';
}

// Applies the current matrix rules as "[MATRIX] key" columns onto the Step 2 data,
// updating existing columns or adding new ones, then re-renders.
function syncMatrixToStep2() {
    const p = projects[activeProjectIdx];
    if (!p.dataA) return;

    const rawInput = document.getElementById('matrixRawInput').value;
    p.rawMatrix = rawInput;
    p.matrix = parseMatrixString(p.rawMatrix);

    if (!p.matrix || p.matrix.length === 0) return showToast("No Matrix Rules found to sync.");

    p.matrix.forEach(rule => {
        let key = rule.key.replace(/^\*/, ''); // drop the universal "*" prefix
        let val = rule.val.trim();
        if (val.toUpperCase() === "NO") val = "FALSE";
        if (val.toUpperCase() === "YES") val = "TRUE";

        let headerName = `[MATRIX] ${key}`;
        
        if (!p.dataA.headers.includes(headerName)) {
            p.dataA.headers.push(headerName);
            p.dataA.body.forEach(row => row.push(val));
            if (p.dataA.hiddenRows) p.dataA.hiddenRows.forEach(hRow => hRow.data.push(val));
        } else {
            let colIdx = p.dataA.headers.indexOf(headerName);
            p.dataA.body.forEach(row => row[colIdx] = val);
            if (p.dataA.hiddenRows) p.dataA.hiddenRows.forEach(hRow => hRow.data[colIdx] = val);
        }
    });

    renderPreviewTables();
    showToast("Matrix Rules Synced into Step 2!");
}


// ======================================================
// ===== START : Split By Columns ========================
// ======================================================

let splitColumnsSide = null;
let splitUniqueColumn = null;

function showSplitColumnsModal(side) {

    splitColumnsSide = side;
    splitUniqueColumn = null;

    const p = projects[activeProjectIdx];
    const data = side === "A" ? p.dataA : p.dataB;

    if (!data || !data.headers) {
        alert("No table loaded.");
        return;
    }

    let html = "";

    // --------------------------------------------------
    // Column checkboxes for Normal Split
    // --------------------------------------------------
    data.headers.forEach((header, index) => {
        html += `
        <label class="splitColumnItem"
               style="display:flex;align-items:center;padding:6px 0;gap:8px;cursor:pointer;">
            <input
                type="checkbox"
                class="split-col"
                value="${index}">
            <span>${header}</span>
        </label>`;
    });

    // --------------------------------------------------
    // Dropdown for Unique Pattern checking
    // --------------------------------------------------
    let uniqueColumnOptions = `<option value="">None (Don't check patterns)</option>`;
    data.headers.forEach((header, index) => {
        uniqueColumnOptions += `<option value="${index}">${header}</option>`;
    });

    const oldOverlay = document.getElementById("splitColumnsOverlay");
    if (oldOverlay) oldOverlay.remove();

    const overlay = document.createElement("div");
    overlay.id = "splitColumnsOverlay";
    overlay.style.cssText = `
        position:fixed;
        inset:0;
        background:rgba(0,0,0,.45);
        display:flex;
        align-items:center;
        justify-content:center;
        z-index:999999;
    `;

    overlay.innerHTML = `
        <div style="
            width:430px;
            max-height:85vh;
            background:white;
            border-radius:8px;
            padding:20px;
            display:flex;
            flex-direction:column;
            box-shadow:0 15px 40px rgba(0,0,0,.20);">

            <h3 style="margin-top:0;margin-bottom:15px;">Split by Columns</h3>

            <!-- SEARCH -->
            <input
                id="splitColumnSearch"
                placeholder="Search columns..."
                onkeyup="filterSplitColumns()"
                style="
                    padding:8px;
                    border:1px solid #ddd;
                    border-radius:4px;
                    margin-bottom:10px;">

            <!-- COLUMN LIST -->
            <div id="splitColumnList"
                 style="
                    max-height:180px;
                    overflow:auto;
                    border:1px solid #eee;
                    padding:10px;">
                ${html}
            </div>

            <!-- UNIQUE PATTERN SUB-SPLIT OPTION -->
            <div style="
                margin-top:15px;
                padding:12px;
                border:1px solid #e2e8f0;
                border-radius:6px;
                background:#f8fafc;">

                <div style="font-size:12px;font-weight:700;color:#334155;margin-bottom:6px;">
                    Check Unique Pattern in Sets
                </div>

                <select
                    id="splitUniqueColumn"
                    onchange="splitUniqueColumnChanged()"
                    style="width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:4px;background:white;font-size:12px;">
                    ${uniqueColumnOptions}
                </select>

                <div style="margin-top:6px;font-size:10px;color:#64748b;line-height:1.4;">
                    If enabled, each column set will be checked for repeating patterns on this column.
                </div>
            </div>

            <!-- BUTTONS -->
            <div style="
                display:flex;
                justify-content:flex-end;
                gap:10px;
                margin-top:15px;">

                <button onclick="closeSplitColumnsModal()">
                    Cancel
                </button>

                <button
                    onclick="splitCurrentTable()"
                    style="
                        background:#2563eb;
                        color:white;
                        border:none;
                        padding:8px 18px;
                        border-radius:4px;
                        cursor:pointer;">
                    Split
                </button>

            </div>

        </div>
    `;

    document.body.appendChild(overlay);
}

function splitUniqueColumnChanged() {
    const select = document.getElementById("splitUniqueColumn");
    if (!select) return;

    if (select.value === "") {
        splitUniqueColumn = null;
    } else {
        splitUniqueColumn = Number(select.value);
    }
}

function closeSplitColumnsModal() {
    const modal = document.getElementById("splitColumnsOverlay");
    if (modal) modal.remove();

    splitColumnsSide = null;
    splitUniqueColumn = null;
}

function filterSplitColumns() {
    const txt = document
        .getElementById("splitColumnSearch")
        .value
        .toLowerCase();

    document
        .querySelectorAll(".splitColumnItem")
        .forEach(item => {
            item.style.display =
                item.innerText.toLowerCase().includes(txt)
                    ? "flex"
                    : "none";
        });
}

// ======================================================
// ===== SPLIT LOGIC (NORMAL + UNIQUE PATTERN IN SETS) ===
// ======================================================

function splitCurrentTable() {

    const currentProject = projects[activeProjectIdx];

    const data = splitColumnsSide === "A"
        ? currentProject.dataA
        : currentProject.dataB;

    if (!data) {
        alert("No table loaded.");
        return;
    }

    // 1. Collect selected columns for primary grouping
    const selectedColumns = [];
    document.querySelectorAll(".split-col:checked").forEach(cb => {
        selectedColumns.push(Number(cb.value));
    });

    if (!selectedColumns.length && (splitUniqueColumn === null || splitUniqueColumn === undefined)) {
        alert("Please select at least one column to split by.");
        return;
    }

    // 2. Perform primary group by selected columns
    let initialGroupedTables = {};

    if (selectedColumns.length > 0) {
        initialGroupedTables = groupTableByColumns(data, selectedColumns);
    } else {
        initialGroupedTables["All Data"] = [[...data.headers], ...data.body];
    }

    // 3. Inspect each set for Unique Patterns (if a column is selected)
    const finalGroupedTables = {};

    if (splitUniqueColumn !== null && splitUniqueColumn !== undefined) {
        const uniqueColIdx = Number(splitUniqueColumn);
        const colName = data.headers[uniqueColIdx] || `Col_${uniqueColIdx}`;

        Object.entries(initialGroupedTables).forEach(([groupKey, rows]) => {

            if (!rows || rows.length <= 1) return;

            // Prepare dataset subset for detectUniquePatterns
            const subData = {
                headers: rows[0],
                body: rows.slice(1)
            };

            const groups = detectUniquePatterns(subData, uniqueColIdx);

            if (groups && groups.length > 1) {
                // Unique pattern detected within this column set -> Sub-split!
                groups.forEach((group, index) => {
                    const rowsForSet = [[...data.headers]];
                    const rowIndices = group.rows || group.rowIndexes || [];

                    rowIndices.forEach(rowIndex => {
                        const idx = Number(rowIndex);
                        let row = subData.body[idx];
                        if (!row && idx > 0 && subData.body[idx - 1]) {
                            row = subData.body[idx - 1];
                        }
                        if (row) {
                            rowsForSet.push([...row]);
                        }
                    });

                    if (rowsForSet.length > 1) {
                        const newKey = selectedColumns.length > 0
                            ? `${groupKey} - ${colName} Pattern ${index + 1}`
                            : `${colName} Pattern ${index + 1}`;
                        finalGroupedTables[newKey] = rowsForSet;
                    }
                });
            } else {
                // No multi-pattern detected; retain standard group set
                finalGroupedTables[groupKey] = rows;
            }
        });
    } else {
        // No pattern check requested -> keep initial column groups as is
        Object.assign(finalGroupedTables, initialGroupedTables);
    }

    // 4. Create project sets using exact naming standard
    const createdIndexes = [];

    Object.entries(finalGroupedTables).forEach(([groupName, rows]) => {

        if (!rows || rows.length <= 1) return;

        createSet(
            `${currentProject.name} - ${groupName}`,
            currentProject.sourceWorkbook,
            currentProject.originalSheetName,
            currentProject.fileName
        );

        const newProject = projects[projects.length - 1];

        if (currentProject.settings) {
            newProject.settings = { ...currentProject.settings };
        }

        if (splitColumnsSide === "A") {
            newProject.rawA = arrayToTSV(rows);
            newProject.dataA = null;
        } else {
            newProject.rawB = arrayToTSV(rows);
            newProject.dataB = null;
        }

        newProject.status = "ready";
        newProject.step = 1;

        createdIndexes.push(projects.length - 1);

    });

    closeSplitColumnsModal();
    refreshProjectTabs();

    if (createdIndexes.length > 0) {
        activeProjectIdx = createdIndexes[0];
        loadProjectIntoView(activeProjectIdx);
    }

    alert(createdIndexes.length + " Sets created successfully.");

}

function groupTableByColumns(data, columnIndexes) {

    const groups = {};

    data.body.forEach(row => {

        const key = columnIndexes
            .map(i => {
                let value = row[i];
                if (value === undefined || value === null || value === "") {
                    value = "(Blank)";
                }
                return String(value).trim();
            })
            .join("-");

        if (!groups[key]) {
            groups[key] = [[...data.headers]];
        }

        groups[key].push([...row]);

    });

    return groups;

}

// ======================================================
// ===== END : Split By Columns ==========================
// ======================================================

// ======================================================
// ===== START : Split By Rows ===========================
// ======================================================

let splitRowsSide = null;
let splitRowGroups = [];

/*
    splitRowGroups structure:

    [
        {
            rows: [0, 1, 2],
            label: "Rows 1-3"
        },
        {
            rows: [5, 6, 7],
            label: "Rows 6-8"
        }
    ]
*/


// ------------------------------------------------------
// Utility: Escape HTML
// ------------------------------------------------------

function escapeSplitRowsHtml(value) {

    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


// ------------------------------------------------------
// Open Split By Rows Modal
// ------------------------------------------------------

function showSplitByRowsModal(side) {

    splitRowsSide = side;
    splitRowGroups = [];

    const p = projects[activeProjectIdx];

    if (!p) {
        alert("No Set is currently selected.");
        return;
    }

    const data = side === "A"
        ? p.dataA
        : p.dataB;

    if (!data || !data.headers || !data.body) {
        alert("No table is available to split.");
        return;
    }

    if (!data.body.length) {
        alert("The table contains no data rows.");
        return;
    }

    // Remove an old modal if one somehow exists.
    const oldModal = document.getElementById("splitRowsOverlay");

    if (oldModal) {
        oldModal.remove();
    }

    const overlay = document.createElement("div");

    overlay.id = "splitRowsOverlay";

    overlay.style.cssText = `
        position: fixed;
        inset: 0;
        background: rgba(15, 23, 42, 0.55);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 999999;
        padding: 20px;
        box-sizing: border-box;
    `;

    overlay.innerHTML = `

        <div style="
            width: 720px;
            max-width: 95vw;
            max-height: 90vh;
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 20px 50px rgba(0,0,0,0.25);
            display: flex;
            flex-direction: column;
            overflow: hidden;
        ">

            <!-- HEADER -->
            <div style="
                padding: 16px 20px;
                border-bottom: 1px solid #e5e7eb;
                display: flex;
                align-items: center;
                justify-content: space-between;
                background: #f8fafc;
            ">

                <div>
                    <div style="
                        font-size: 16px;
                        font-weight: 700;
                        color: #0f172a;
                    ">
                        <i class="fas fa-cut"
                           style="color:#2563eb;margin-right:7px;"></i>
                        Split by Rows
                    </div>

                    <div style="
                        margin-top: 3px;
                        font-size: 11px;
                        color: #64748b;
                    ">
                        Select rows or enter ranges to create separate Sets
                    </div>
                </div>

                <button
                    onclick="closeSplitByRowsModal()"
                    style="
                        width: 32px;
                        height: 32px;
                        border: none;
                        background: transparent;
                        color: #64748b;
                        font-size: 18px;
                        cursor: pointer;
                        border-radius: 6px;
                    "
                    onmouseover="this.style.background='#e2e8f0'"
                    onmouseout="this.style.background='transparent'"
                >
                    <i class="fas fa-times"></i>
                </button>

            </div>


            <!-- BODY -->
            <div style="
                padding: 16px 20px;
                overflow: auto;
                flex: 1;
            ">


                <!-- RANGE SECTION -->
                <div style="
                    border: 1px solid #e2e8f0;
                    border-radius: 8px;
                    padding: 12px;
                    background: #f8fafc;
                    margin-bottom: 14px;
                ">

                    <div style="
                        font-size: 12px;
                        font-weight: 700;
                        color: #334155;
                        margin-bottom: 8px;
                    ">
                        <i class="fas fa-arrows-alt-h"
                           style="color:#2563eb;margin-right:5px;"></i>
                        Add rows by range
                    </div>

                    <div style="
                        display:flex;
                        gap:8px;
                        align-items:center;
                    ">

                        <input
                            id="splitRowRangeInput"
                            type="text"
                            placeholder="Example: 1-10"
                            style="
                                flex:1;
                                padding:8px 10px;
                                border:1px solid #cbd5e1;
                                border-radius:6px;
                                outline:none;
                                font-size:12px;
                                box-sizing:border-box;
                            "
                        >

                        <button
                            onclick="addSplitRowRangeGroup()"
                            style="
                                background:#2563eb;
                                color:white;
                                border:none;
                                padding:8px 13px;
                                border-radius:6px;
                                font-size:12px;
                                font-weight:600;
                                cursor:pointer;
                            "
                        >
                            <i class="fas fa-plus"></i>
                            Add Range
                        </button>

                    </div>

                    <div style="
                        margin-top:6px;
                        font-size:10px;
                        color:#94a3b8;
                    ">
                        Examples: 1-10 &nbsp;|&nbsp; 5-15 &nbsp;|&nbsp; 20
                    </div>

                </div>


                <!-- DIRECT SELECTION SECTION -->
                <div style="
                    border:1px solid #e2e8f0;
                    border-radius:8px;
                    overflow:hidden;
                    margin-bottom:14px;
                ">

                    <div style="
                        padding:10px 12px;
                        background:#f8fafc;
                        border-bottom:1px solid #e2e8f0;
                        display:flex;
                        justify-content:space-between;
                        align-items:center;
                    ">

                        <div>
                            <div style="
                                font-size:12px;
                                font-weight:700;
                                color:#334155;
                            ">
                                <i class="fas fa-mouse-pointer"
                                   style="color:#2563eb;margin-right:5px;"></i>
                                Select rows directly
                            </div>

                            <div style="
                                font-size:10px;
                                color:#94a3b8;
                                margin-top:2px;
                            ">
                                Select any rows and add them as one group
                            </div>
                        </div>

                        <span
                            id="splitSelectedRowCount"
                            style="
                                font-size:11px;
                                color:#2563eb;
                                font-weight:700;
                            "
                        >
                            0 selected
                        </span>

                    </div>


                    <!-- ROW LIST -->
                    <div
                        id="splitRowsList"
                        style="
                            max-height:260px;
                            overflow:auto;
                            padding:5px;
                        "
                    >

                        ${buildSplitRowsList(data)}

                    </div>


                    <!-- SELECTION ACTION -->
                    <div style="
                        padding:9px 12px;
                        border-top:1px solid #e2e8f0;
                        background:#fafafa;
                        display:flex;
                        justify-content:flex-end;
                    ">

                        <button
                            onclick="addSelectedRowsGroup()"
                            style="
                                background:#0f766e;
                                color:white;
                                border:none;
                                padding:7px 12px;
                                border-radius:6px;
                                font-size:11px;
                                font-weight:600;
                                cursor:pointer;
                            "
                        >
                            <i class="fas fa-layer-group"></i>
                            Add Selected Rows as Group
                        </button>

                    </div>

                </div>


                <!-- GROUPS -->
                <div style="
                    border:1px solid #e2e8f0;
                    border-radius:8px;
                    overflow:hidden;
                ">

                    <div style="
                        padding:10px 12px;
                        background:#f8fafc;
                        border-bottom:1px solid #e2e8f0;
                        display:flex;
                        justify-content:space-between;
                        align-items:center;
                    ">

                        <div style="
                            font-size:12px;
                            font-weight:700;
                            color:#334155;
                        ">
                            <i class="fas fa-layer-group"
                               style="color:#7c3aed;margin-right:5px;"></i>
                            Groups
                        </div>

                        <span
                            id="splitRowGroupCount"
                            style="
                                font-size:10px;
                                color:#64748b;
                            "
                        >
                            0 groups
                        </span>

                    </div>


                    <div
                        id="splitRowGroupsList"
                        style="
                            min-height:70px;
                            max-height:190px;
                            overflow:auto;
                            padding:8px;
                        "
                    >

                        <div style="
                            text-align:center;
                            padding:20px;
                            color:#94a3b8;
                            font-size:11px;
                        ">
                            No groups added yet.
                        </div>

                    </div>

                </div>


                <!-- REMAINING ROWS -->
                <label style="
                    display:flex;
                    align-items:center;
                    gap:8px;
                    margin-top:12px;
                    padding:10px 12px;
                    border:1px solid #e2e8f0;
                    border-radius:7px;
                    background:#ffffff;
                    cursor:pointer;
                ">

                    <input
                        type="checkbox"
                        id="splitRemainingRows"
                        style="cursor:pointer;"
                    >

                    <div>
                        <div style="
                            font-size:11px;
                            font-weight:600;
                            color:#334155;
                        ">
                            Put remaining rows into another Set
                        </div>

                        <div style="
                            font-size:10px;
                            color:#94a3b8;
                            margin-top:2px;
                        ">
                            Rows not assigned to any group will be included.
                        </div>
                    </div>

                </label>

            </div>


            <!-- FOOTER -->
            <div style="
                padding:12px 20px;
                border-top:1px solid #e5e7eb;
                display:flex;
                justify-content:space-between;
                align-items:center;
                background:#f8fafc;
            ">

                <div
                    id="splitRowsFooterInfo"
                    style="
                        font-size:10px;
                        color:#64748b;
                    "
                >
                    No groups created
                </div>

                <div style="
                    display:flex;
                    gap:8px;
                ">

                    <button
                        onclick="closeSplitByRowsModal()"
                        style="
                            background:white;
                            color:#475569;
                            border:1px solid #cbd5e1;
                            padding:8px 14px;
                            border-radius:6px;
                            font-size:11px;
                            font-weight:600;
                            cursor:pointer;
                        "
                    >
                        Cancel
                    </button>

                    <button
                        onclick="createSetsFromRowGroups()"
                        id="createRowSplitSetsBtn"
                        disabled
                        style="
                            background:#2563eb;
                            color:white;
                            border:none;
                            padding:8px 15px;
                            border-radius:6px;
                            font-size:11px;
                            font-weight:600;
                            cursor:not-allowed;
                            opacity:.5;
                        "
                    >
                        <i class="fas fa-code-branch"></i>
                        Create Sets
                    </button>

                </div>

            </div>

        </div>
    `;

    document.body.appendChild(overlay);

    updateSplitRowsUI();

}


// ------------------------------------------------------
// Build Row List
// ------------------------------------------------------

function buildSplitRowsList(data) {

    let html = "";

    data.body.forEach((row, index) => {

        const rowNumber = index + 1;

        const previewValues = row
            .slice(0, 4)
            .map(cell => String(cell ?? "").trim())
            .filter(v => v !== "")
            .join(" | ");

        html += `
            <label
                class="splitRowItem"
                data-row-index="${index}"
                style="
                    display:flex;
                    align-items:center;
                    gap:8px;
                    padding:7px 8px;
                    border-radius:5px;
                    cursor:pointer;
                    border-bottom:1px solid #f1f5f9;
                    font-size:11px;
                "
                onmouseover="this.style.background='#f8fafc'"
                onmouseout="this.style.background='transparent'"
            >

                <input
                    type="checkbox"
                    class="split-row-checkbox"
                    value="${index}"
                    onchange="updateSplitRowsUI()"
                    style="cursor:pointer;"
                >

                <span style="
                    width:38px;
                    color:#64748b;
                    font-weight:700;
                    text-align:right;
                ">
                    ${rowNumber}
                </span>

                <span style="
                    flex:1;
                    color:#334155;
                    white-space:nowrap;
                    overflow:hidden;
                    text-overflow:ellipsis;
                "
                title="${escapeSplitRowsHtml(previewValues)}">
                    ${escapeSplitRowsHtml(previewValues || "(empty row)")}
                </span>

            </label>
        `;

    });

    return html;
}


// ------------------------------------------------------
// Update Selected Count / Groups UI
// ------------------------------------------------------

function updateSplitRowsUI() {

    const selected = Array.from(
        document.querySelectorAll(".split-row-checkbox:checked")
    ).map(cb => Number(cb.value));

    const selectedCount =
        document.getElementById("splitSelectedRowCount");

    if (selectedCount) {

        selectedCount.innerText =
            `${selected.length} selected`;

    }

    renderSplitRowGroups();

}


// ------------------------------------------------------
// Get Rows Already Assigned To Groups
// ------------------------------------------------------

function getAssignedSplitRows() {

    const assigned = new Set();

    splitRowGroups.forEach(group => {

        group.rows.forEach(rowIndex => {

            assigned.add(rowIndex);

        });

    });

    return assigned;
}


// ------------------------------------------------------
// Add Manually Selected Rows As Group
// ------------------------------------------------------

function addSelectedRowsGroup() {

    const checkboxes =
        document.querySelectorAll(
            ".split-row-checkbox:checked"
        );

    const selectedRows = Array.from(checkboxes)
        .map(cb => Number(cb.value))
        .sort((a, b) => a - b);

    if (!selectedRows.length) {

        alert("Please select at least one row.");

        return;

    }

    const assigned = getAssignedSplitRows();

    const duplicateRows = selectedRows.filter(
        row => assigned.has(row)
    );

    if (duplicateRows.length) {

        alert(
            "Some selected rows are already assigned to another group:\n\n" +
            duplicateRows
                .map(row => row + 1)
                .join(", ")
        );

        return;

    }

    splitRowGroups.push({

        rows: selectedRows,

        label: buildSplitRowGroupLabel(
            selectedRows
        )

    });

    // Clear current selection.
    document
        .querySelectorAll(".split-row-checkbox:checked")
        .forEach(cb => {

            cb.checked = false;

        });

    updateSplitRowsUI();

}


// ------------------------------------------------------
// Add Range Group
// ------------------------------------------------------

function addSplitRowRangeGroup() {

    const input =
        document.getElementById("splitRowRangeInput");

    if (!input) return;

    const value = input.value.trim();

    if (!value) {

        alert("Enter a row range, for example 1-10.");

        input.focus();

        return;

    }

    const range = parseSplitRowRange(value);

    if (!range) {

        alert(
            "Invalid row range.\n\n" +
            "Use formats like:\n" +
            "1-10\n" +
            "5-15\n" +
            "20"
        );

        input.focus();

        return;

    }

    const p = projects[activeProjectIdx];

    const data = splitRowsSide === "A"
        ? p.dataA
        : p.dataB;

    if (!data || !data.body) {

        alert("No table available.");

        return;

    }

    const start = range.start;
    const end = range.end;

    if (
        start < 0 ||
        end >= data.body.length
    ) {

        alert(
            `Row range must be between 1 and ${data.body.length}.`
        );

        return;

    }

    const rows = [];

    for (
        let i = start;
        i <= end;
        i++
    ) {

        rows.push(i);

    }

    const assigned = getAssignedSplitRows();

    const duplicateRows = rows.filter(
        row => assigned.has(row)
    );

    if (duplicateRows.length) {

        alert(
            "Some rows are already assigned to another group:\n\n" +
            duplicateRows
                .map(row => row + 1)
                .join(", ")
        );

        return;

    }

    splitRowGroups.push({

        rows: rows,

        label:
            start === end
                ? `Row ${start + 1}`
                : `Rows ${start + 1}-${end + 1}`

    });

    input.value = "";

    updateSplitRowsUI();

}


// ------------------------------------------------------
// Parse Range
// ------------------------------------------------------

function parseSplitRowRange(value) {

    const cleaned = String(value)
        .trim()
        .replace(/\s+/g, "");

    // Single row: "10"
    if (/^\d+$/.test(cleaned)) {

        const row = Number(cleaned);

        if (row < 1) return null;

        return {
            start: row - 1,
            end: row - 1
        };

    }

    // Range: "10-20"
    const match = cleaned.match(/^(\d+)-(\d+)$/);

    if (!match) return null;

    const first = Number(match[1]);
    const second = Number(match[2]);

    if (
        first < 1 ||
        second < 1
    ) {
        return null;
    }

    const start = Math.min(first, second);
    const end = Math.max(first, second);

    return {
        start: start - 1,
        end: end - 1
    };

}


// ------------------------------------------------------
// Group Label
// ------------------------------------------------------

function buildSplitRowGroupLabel(rows) {

    if (!rows.length) {
        return "Empty Group";
    }

    if (rows.length === 1) {

        return `Row ${rows[0] + 1}`;

    }

    // Check if consecutive.
    let consecutive = true;

    for (let i = 1; i < rows.length; i++) {

        if (rows[i] !== rows[i - 1] + 1) {

            consecutive = false;

            break;

        }

    }

    if (consecutive) {

        return `Rows ${rows[0] + 1}-${rows[rows.length - 1] + 1}`;

    }

    return `Rows ${rows.map(r => r + 1).join(", ")}`;

}


// ------------------------------------------------------
// Render Groups
// ------------------------------------------------------

function renderSplitRowGroups() {

    const container =
        document.getElementById(
            "splitRowGroupsList"
        );

    const count =
        document.getElementById(
            "splitRowGroupCount"
        );

    const footer =
        document.getElementById(
            "splitRowsFooterInfo"
        );

    const createBtn =
        document.getElementById(
            "createRowSplitSetsBtn"
        );

    if (!container) return;

    if (count) {

        count.innerText =
            `${splitRowGroups.length} ${
                splitRowGroups.length === 1
                    ? "group"
                    : "groups"
            }`;

    }

    if (!splitRowGroups.length) {

        container.innerHTML = `
            <div style="
                text-align:center;
                padding:20px;
                color:#94a3b8;
                font-size:11px;
            ">
                No groups added yet.
            </div>
        `;

        if (footer) {

            footer.innerText =
                "Add rows or ranges to create groups";

        }

        if (createBtn) {

            createBtn.disabled = true;
            createBtn.style.opacity = ".5";
            createBtn.style.cursor = "not-allowed";

        }

        return;

    }

    let html = "";

    let totalRows = 0;

    splitRowGroups.forEach((group, index) => {

        totalRows += group.rows.length;

        html += `

            <div style="
                display:flex;
                align-items:center;
                justify-content:space-between;
                gap:10px;
                padding:9px 10px;
                border:1px solid #e2e8f0;
                border-radius:6px;
                margin-bottom:6px;
                background:#ffffff;
            ">

                <div style="
                    display:flex;
                    align-items:center;
                    gap:9px;
                    min-width:0;
                ">

                    <div style="
                        width:28px;
                        height:28px;
                        border-radius:6px;
                        background:#ede9fe;
                        color:#7c3aed;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        font-size:11px;
                        font-weight:700;
                        flex-shrink:0;
                    ">
                        ${index + 1}
                    </div>

                    <div style="min-width:0;">

                        <div style="
                            font-size:11px;
                            font-weight:700;
                            color:#334155;
                        ">
                            Group ${index + 1}
                        </div>

                        <div style="
                            font-size:10px;
                            color:#64748b;
                            margin-top:2px;
                            white-space:nowrap;
                            overflow:hidden;
                            text-overflow:ellipsis;
                        ">
                            ${escapeSplitRowsHtml(group.label)}
                            &nbsp; • &nbsp;
                            ${group.rows.length} ${
                                group.rows.length === 1
                                    ? "row"
                                    : "rows"
                            }
                        </div>

                    </div>

                </div>

                <button
                    onclick="removeSplitRowGroup(${index})"
                    style="
                        width:27px;
                        height:27px;
                        border:1px solid #fecaca;
                        background:#fff1f2;
                        color:#ef4444;
                        border-radius:5px;
                        cursor:pointer;
                        flex-shrink:0;
                    "
                    title="Remove group"
                >
                    <i class="fas fa-times"></i>
                </button>

            </div>

        `;

    });

    container.innerHTML = html;

    if (footer) {

        footer.innerText =
            `${splitRowGroups.length} ${
                splitRowGroups.length === 1
                    ? "group"
                    : "groups"
            } • ${totalRows} rows assigned`;

    }

    if (createBtn) {

        createBtn.disabled = false;
        createBtn.style.opacity = "1";
        createBtn.style.cursor = "pointer";

    }

}


// ------------------------------------------------------
// Remove Group
// ------------------------------------------------------

function removeSplitRowGroup(index) {

    if (
        index < 0 ||
        index >= splitRowGroups.length
    ) {
        return;
    }

    splitRowGroups.splice(index, 1);

    updateSplitRowsUI();

}


// ------------------------------------------------------
// Create Sets From Groups
// ------------------------------------------------------

function createSetsFromRowGroups() {

    if (!splitRowGroups.length) {

        alert("Please create at least one row group.");

        return;

    }

    const currentProject =
        projects[activeProjectIdx];

    if (!currentProject) {

        alert("No active Set found.");

        return;

    }

    const data =
        splitRowsSide === "A"
            ? currentProject.dataA
            : currentProject.dataB;

    if (!data || !data.body) {

        alert("No table data found.");

        return;

    }

    // ----------------------------------------------
    // Ask whether remaining rows should be included
    // ----------------------------------------------

    const remainingCheckbox =
        document.getElementById(
            "splitRemainingRows"
        );

    const includeRemaining =
        remainingCheckbox
            ? remainingCheckbox.checked
            : false;

    const assigned =
        getAssignedSplitRows();

    const remainingRows = [];

    if (includeRemaining) {

        for (
            let i = 0;
            i < data.body.length;
            i++
        ) {

            if (!assigned.has(i)) {

                remainingRows.push(i);

            }

        }

        if (remainingRows.length > 0) {

            splitRowGroups.push({

                rows: remainingRows,

                label:
                    remainingRows.length === 1
                        ? `Remaining Row ${remainingRows[0] + 1}`
                        : `Remaining Rows (${remainingRows.length})`

            });

        }

    }


    if (!splitRowGroups.length) {

        alert("No rows available for splitting.");

        return;

    }


    // ----------------------------------------------
    // Create all new Sets
    // ----------------------------------------------

    const createdIndexes = [];

    splitRowGroups.forEach((group, groupIndex) => {

        if (!group.rows.length) {
            return;
        }

        const rowsForSet = [
            [...data.headers]
        ];

        group.rows.forEach(rowIndex => {

            if (
                data.body[rowIndex]
            ) {

                rowsForSet.push([
                    ...data.body[rowIndex]
                ]);

            }

        });

        if (rowsForSet.length <= 1) {
            return;
        }


        const setName =
            `${currentProject.name} - Row Split ${groupIndex + 1}`;


        // Use your existing createSet()
        createSet(
            setName,
            currentProject.sourceWorkbook,
            currentProject.originalSheetName,
            currentProject.fileName
        );


        const newProject =
            projects[projects.length - 1];


        // ------------------------------------------
        // Preserve useful settings from original Set
        // ------------------------------------------

        if (currentProject.settings) {

            newProject.settings = {
                ...currentProject.settings
            };

        }


        // ------------------------------------------
        // Put split data into the correct side
        // ------------------------------------------

        if (splitRowsSide === "A") {

            newProject.rawA =
                arrayToTSV(rowsForSet);

            // Important:
            // Leave dataA null.
            // jumpToStep(2) will parse it using
            // your existing parseExcelData() flow.

            newProject.dataA = null;

        } else {

            newProject.rawB =
                arrayToTSV(rowsForSet);

            // Important:
            // Leave dataB null.
            // jumpToStep(2) will parse it using
            // your existing parseExcelData() flow.

            newProject.dataB = null;

        }


        // ------------------------------------------
        // IMPORTANT:
        // New split Sets must start at Step 1
        // ------------------------------------------

        newProject.status = "ready";
        newProject.step = 1;


        createdIndexes.push(
            projects.length - 1
        );

    });


    // ----------------------------------------------
    // Clean modal state
    // ----------------------------------------------

    splitRowGroups = [];

    closeSplitByRowsModal();


    // ----------------------------------------------
    // Refresh Set tabs/list
    // ----------------------------------------------

    if (
        typeof refreshProjectTabs === "function"
    ) {

        refreshProjectTabs();

    }


    // ----------------------------------------------
    // Open first newly created Set
    // ----------------------------------------------

    if (createdIndexes.length > 0) {

        activeProjectIdx =
            createdIndexes[0];

        // loadProjectIntoView() uses:
        //
        // jumpToStep(p.step || 1)
        //
        // and because we explicitly set:
        //
        // newProject.step = 1
        //
        // the new Set opens at Step 1.

        loadProjectIntoView(
            activeProjectIdx
        );

    }


    alert(
        `${createdIndexes.length} Sets created successfully.`
    );

}


// ------------------------------------------------------
// Close Modal
// ------------------------------------------------------

function closeSplitByRowsModal() {

    const modal =
        document.getElementById(
            "splitRowsOverlay"
        );

    if (modal) {

        modal.remove();

    }

    splitRowGroups = [];
    splitRowsSide = null;

}


// ======================================================
// ===== END : Split By Rows =============================
// ======================================================

// ======================================================
// ===== START : Split By Unique Pattern =================
// ======================================================

let uniquePatternSide = null;


// ------------------------------------------------------
// Open Unique Pattern Modal
// ------------------------------------------------------

function showSplitByUniquePatternModal(side) {

    uniquePatternSide = side;

    const p = projects[activeProjectIdx];

    if (!p) {
        alert("No active Set found.");
        return;
    }

    const data = side === "A"
        ? p.dataA
        : p.dataB;

    if (!data || !data.headers || !data.body) {
        alert("No table available.");
        return;
    }

    if (!data.body.length) {
        alert("The table contains no data rows.");
        return;
    }

    const oldModal =
        document.getElementById("uniquePatternOverlay");

    if (oldModal) {
        oldModal.remove();
    }

    let columnOptions = "";

    data.headers.forEach((header, index) => {

        columnOptions += `
            <option value="${index}">
                ${escapeUniquePatternHtml(header)}
            </option>
        `;

    });

    const overlay = document.createElement("div");

    overlay.id = "uniquePatternOverlay";

    overlay.style.cssText = `
        position:fixed;
        inset:0;
        background:rgba(15,23,42,.55);
        display:flex;
        align-items:center;
        justify-content:center;
        z-index:999999;
        padding:20px;
        box-sizing:border-box;
    `;

    overlay.innerHTML = `

        <div style="
            width:650px;
            max-width:95vw;
            max-height:90vh;
            background:#fff;
            border-radius:12px;
            box-shadow:0 20px 50px rgba(0,0,0,.25);
            display:flex;
            flex-direction:column;
            overflow:hidden;
        ">

            <!-- HEADER -->
            <div style="
                padding:16px 20px;
                border-bottom:1px solid #e5e7eb;
                background:#f8fafc;
                display:flex;
                align-items:center;
                justify-content:space-between;
            ">

                <div>

                    <div style="
                        font-size:16px;
                        font-weight:700;
                        color:#0f172a;
                    ">
                        <i class="fas fa-layer-group"
                           style="color:#7c3aed;margin-right:7px;"></i>
                        Split by Unique Pattern
                    </div>

                    <div style="
                        margin-top:4px;
                        font-size:11px;
                        color:#64748b;
                    ">
                        Automatically split rows into repeated value patterns
                    </div>

                </div>

                <button
                    onclick="closeSplitByUniquePatternModal()"
                    style="
                        width:32px;
                        height:32px;
                        border:none;
                        background:transparent;
                        color:#64748b;
                        font-size:18px;
                        cursor:pointer;
                        border-radius:6px;
                    "
                >
                    <i class="fas fa-times"></i>
                </button>

            </div>


            <!-- BODY -->
            <div style="
                padding:18px 20px;
                overflow:auto;
                flex:1;
            ">


                <!-- COLUMN SELECT -->
                <div style="
                    border:1px solid #e2e8f0;
                    border-radius:8px;
                    padding:14px;
                    background:#f8fafc;
                    margin-bottom:14px;
                ">

                    <div style="
                        font-size:12px;
                        font-weight:700;
                        color:#334155;
                        margin-bottom:8px;
                    ">
                        Select column
                    </div>

                    <select
                        id="uniquePatternColumnSelect"
                        onchange="previewUniquePatternSplit()"
                        style="
                            width:100%;
                            padding:9px 10px;
                            border:1px solid #cbd5e1;
                            border-radius:6px;
                            background:white;
                            font-size:12px;
                            color:#334155;
                            outline:none;
                            cursor:pointer;
                        "
                    >
                        ${columnOptions}
                    </select>

                    <div style="
                        margin-top:7px;
                        font-size:10px;
                        color:#94a3b8;
                    ">
                        Example: SIZE → A B C D A B
                        becomes A B C D + A B
                    </div>

                </div>


                <!-- PREVIEW -->
                <div style="
                    border:1px solid #e2e8f0;
                    border-radius:8px;
                    overflow:hidden;
                ">

                    <div style="
                        padding:10px 12px;
                        background:#f8fafc;
                        border-bottom:1px solid #e2e8f0;
                        display:flex;
                        justify-content:space-between;
                        align-items:center;
                    ">

                        <div style="
                            font-size:12px;
                            font-weight:700;
                            color:#334155;
                        ">
                            Detected Patterns
                        </div>

                        <span
                            id="uniquePatternSetCount"
                            style="
                                font-size:10px;
                                color:#7c3aed;
                                font-weight:700;
                            "
                        >
                            0 Sets
                        </span>

                    </div>


                    <div
                        id="uniquePatternPreview"
                        style="
                            max-height:300px;
                            overflow:auto;
                            padding:10px;
                        "
                    >
                        <!-- generated -->
                    </div>

                </div>


                <!-- INFO -->
                <div
                    id="uniquePatternInfo"
                    style="
                        margin-top:12px;
                        padding:10px 12px;
                        border-radius:7px;
                        background:#f5f3ff;
                        border:1px solid #ddd6fe;
                        color:#6d28d9;
                        font-size:10px;
                        line-height:1.5;
                    "
                >
                    Select a column to detect patterns.
                </div>

            </div>


            <!-- FOOTER -->
            <div style="
                padding:12px 20px;
                border-top:1px solid #e5e7eb;
                background:#f8fafc;
                display:flex;
                justify-content:flex-end;
                gap:8px;
            ">

                <button
                    onclick="closeSplitByUniquePatternModal()"
                    style="
                        background:white;
                        color:#475569;
                        border:1px solid #cbd5e1;
                        padding:8px 14px;
                        border-radius:6px;
                        font-size:11px;
                        font-weight:600;
                        cursor:pointer;
                    "
                >
                    Cancel
                </button>

                <button
                    id="createUniquePatternSetsBtn"
                    onclick="createSetsFromUniquePatterns()"
                    style="
                        background:#7c3aed;
                        color:white;
                        border:none;
                        padding:8px 15px;
                        border-radius:6px;
                        font-size:11px;
                        font-weight:600;
                        cursor:pointer;
                    "
                >
                    <i class="fas fa-code-branch"></i>
                    Create Sets
                </button>

            </div>

        </div>
    `;

    document.body.appendChild(overlay);

    previewUniquePatternSplit();
}


// ------------------------------------------------------
// HTML Escape
// ------------------------------------------------------

function escapeUniquePatternHtml(value) {

    if (value === null || value === undefined) {
        return "";
    }

    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}


// ------------------------------------------------------
// Detect Unique Patterns
//
// Example:
//
// A B C D A B
//
// Result:
//
// Group 1 -> A B C D
// Group 2 -> A B
//
// The first repeated value starts a new pattern.
// ------------------------------------------------------

function detectUniquePatterns(data, columnIndex) {

    if (!data || !data.body) {
        return [];
    }

    const groups = [];

    let currentGroup = [];
    let currentValues = new Set();

    data.body.forEach((row, rowIndex) => {

        const rawValue =
            row[columnIndex];

        const value =
            rawValue === null ||
            rawValue === undefined
                ? ""
                : String(rawValue).trim();

        /*
            Empty values are treated as normal values,
            but consecutive empty values stay together.
        */

        /*
            If this value already exists in the
            current pattern, it means the current
            pattern has completed and a new pattern
            should start here.
        */

        if (
            currentGroup.length > 0 &&
            currentValues.has(value)
        ) {

            groups.push({
                rows: currentGroup,
                values: Array.from(currentValues)
            });

            currentGroup = [];
            currentValues = new Set();

        }

        currentGroup.push(rowIndex);
        currentValues.add(value);

    });


    // Push final group
    if (currentGroup.length > 0) {

        groups.push({
            rows: currentGroup,
            values: Array.from(currentValues)
        });

    }

    return groups;
}


// ------------------------------------------------------
// Preview Patterns
// ------------------------------------------------------

function previewUniquePatternSplit() {

    const select =
        document.getElementById(
            "uniquePatternColumnSelect"
        );

    const preview =
        document.getElementById(
            "uniquePatternPreview"
        );

    const count =
        document.getElementById(
            "uniquePatternSetCount"
        );

    const info =
        document.getElementById(
            "uniquePatternInfo"
        );

    if (!select || !preview) {
        return;
    }

    const p =
        projects[activeProjectIdx];

    if (!p) return;

    const data =
        uniquePatternSide === "A"
            ? p.dataA
            : p.dataB;

    if (!data) return;

    const columnIndex =
        Number(select.value);

    const columnName =
        data.headers[columnIndex];

    const groups =
        detectUniquePatterns(
            data,
            columnIndex
        );


    // ----------------------------------------------
    // Set count
    // ----------------------------------------------

    if (count) {

        count.innerText =
            `${groups.length} ${
                groups.length === 1
                    ? "Set"
                    : "Sets"
            }`;

    }


    // ----------------------------------------------
    // Empty result
    // ----------------------------------------------

    if (!groups.length) {

        preview.innerHTML = `
            <div style="
                text-align:center;
                padding:25px;
                color:#94a3b8;
                font-size:11px;
            ">
                No patterns detected.
            </div>
        `;

        return;

    }


    // ----------------------------------------------
    // Render groups
    // ----------------------------------------------

    let html = "";

    groups.forEach((group, index) => {

        const values =
            group.values
                .map(v =>
                    v === ""
                        ? "(blank)"
                        : v
                );

        html += `

            <div style="
                border:1px solid #e2e8f0;
                border-radius:7px;
                padding:10px;
                margin-bottom:7px;
                background:#fff;
            ">

                <div style="
                    display:flex;
                    align-items:center;
                    justify-content:space-between;
                    margin-bottom:7px;
                ">

                    <div style="
                        display:flex;
                        align-items:center;
                        gap:7px;
                    ">

                        <span style="
                            width:25px;
                            height:25px;
                            border-radius:5px;
                            background:#ede9fe;
                            color:#7c3aed;
                            display:flex;
                            align-items:center;
                            justify-content:center;
                            font-size:10px;
                            font-weight:700;
                        ">
                            ${index + 1}
                        </span>

                        <span style="
                            font-size:11px;
                            font-weight:700;
                            color:#334155;
                        ">
                            Set ${index + 1}
                        </span>

                    </div>

                    <span style="
                        font-size:10px;
                        color:#64748b;
                    ">
                        ${group.rows.length} rows
                    </span>

                </div>


                <div style="
                    display:flex;
                    flex-wrap:wrap;
                    gap:5px;
                ">

                    ${values.map(value => `
                        <span style="
                            display:inline-block;
                            padding:4px 7px;
                            background:#f8fafc;
                            border:1px solid #e2e8f0;
                            border-radius:4px;
                            color:#475569;
                            font-size:10px;
                            font-weight:600;
                        ">
                            ${escapeUniquePatternHtml(value)}
                        </span>
                    `).join("")}

                </div>

            </div>
        `;

    });

    preview.innerHTML = html;


    // ----------------------------------------------
    // Information
    // ----------------------------------------------

    if (info) {

        info.innerHTML = `
            <strong>${escapeUniquePatternHtml(columnName)}</strong>
            contains
            <strong>${groups.length}</strong>
            detected pattern${
                groups.length === 1 ? "" : "s"
            }.

            The first repeated value starts a new pattern.
        `;

    }

}


// ------------------------------------------------------
// Create Sets
// ------------------------------------------------------

function createSetsFromUniquePatterns() {

    const select =
        document.getElementById(
            "uniquePatternColumnSelect"
        );

    if (!select) {
        return;
    }

    const p =
        projects[activeProjectIdx];

    if (!p) {

        alert("No active Set found.");

        return;

    }

    const data =
        uniquePatternSide === "A"
            ? p.dataA
            : p.dataB;

    if (!data || !data.body) {

        alert("No table data found.");

        return;

    }

    const columnIndex =
        Number(select.value);

    const columnName =
        data.headers[columnIndex];

    const groups =
        detectUniquePatterns(
            data,
            columnIndex
        );

    if (!groups.length) {

        alert("No patterns were detected.");

        return;

    }


    // ----------------------------------------------
    // Create Sets
    // ----------------------------------------------

    const createdIndexes = [];

    groups.forEach((group, groupIndex) => {

        const rowsForSet = [
            [...data.headers]
        ];


        group.rows.forEach(rowIndex => {

            if (data.body[rowIndex]) {

                rowsForSet.push([
                    ...data.body[rowIndex]
                ]);

            }

        });


        if (rowsForSet.length <= 1) {
            return;
        }


        const setName =
            `${p.name} - ${columnName} Pattern ${groupIndex + 1}`;


        // Use existing createSet()
        createSet(
            setName,
            p.sourceWorkbook,
            p.originalSheetName,
            p.fileName
        );


        const newProject =
            projects[projects.length - 1];


        // ------------------------------------------
        // Preserve settings
        // ------------------------------------------

        if (p.settings) {

            newProject.settings = {
                ...p.settings
            };

        }


        // ------------------------------------------
        // Put data into same side
        // ------------------------------------------

        if (uniquePatternSide === "A") {

            newProject.rawA =
                arrayToTSV(rowsForSet);

            newProject.dataA = null;

        } else {

            newProject.rawB =
                arrayToTSV(rowsForSet);

            newProject.dataB = null;

        }


        // ------------------------------------------
        // IMPORTANT:
        // Always open new Set at Step 1
        // ------------------------------------------

        newProject.step = 1;
        newProject.status = "ready";


        createdIndexes.push(
            projects.length - 1
        );

    });


    // ----------------------------------------------
    // Close modal
    // ----------------------------------------------

    closeSplitByUniquePatternModal();


    // ----------------------------------------------
    // Refresh project list
    // ----------------------------------------------

    if (
        typeof refreshProjectTabs === "function"
    ) {

        refreshProjectTabs();

    }


    // ----------------------------------------------
    // Open first newly created Set
    // ----------------------------------------------

    if (createdIndexes.length > 0) {

        activeProjectIdx =
            createdIndexes[0];

        loadProjectIntoView(
            activeProjectIdx
        );

    }


    alert(
        `${createdIndexes.length} Sets created using ${columnName} patterns.`
    );

}


// ------------------------------------------------------
// Close Modal
// ------------------------------------------------------

function closeSplitByUniquePatternModal() {

    const modal =
        document.getElementById(
            "uniquePatternOverlay"
        );

    if (modal) {
        modal.remove();
    }

    uniquePatternSide = null;

}


// ======================================================
// ===== END : Split By Unique Pattern ===================
// ======================================================