// API Client for ExcelDataEntryWeb with Full In-Browser Client-Side Engine for GitHub Pages
const API_BASE = '/api';

// Detect if running in pure client-side environment (like GitHub Pages) or local backend
const isClientOnlyMode = window.location.hostname.includes('github.io') || 
                         window.location.protocol === 'file:' || 
                         window.USE_CLIENT_ENGINE === true;

// ─── PURE JAVASCRIPT IN-BROWSER WORKBOOK ENGINE ─────────────────────────────
class ClientWorkbookEngine {
    constructor() {
        this.workbook = null;
        this.fileName = 'data.xlsx';
        this.fileHandle = null;
        this.selectedSheet = '';
        this.headerRow = 3;
        this.nameColumn = 2;
        this.sampleRow = 4;
        this.autoSkipBlank = true;
        this.suggestionsDisabled = false;
        this.isAutoSave = true;
        this.hiddenHeaders = new Set();
        this.showAllHeaders = true;
        this.records = [];
        this.isLoaded = false;
    }

    async openFromFile(file) {
        if (!file) throw new Error('Không có file để mở.');
        this.fileName = file.name || 'data.xlsx';
        const buffer = await file.arrayBuffer();
        return this.openFromBuffer(buffer, this.fileName);
    }

    openFromBuffer(buffer, fileName = 'data.xlsx') {
        if (!window.XLSX) {
            throw new Error('Thư viện SheetJS (XLSX) chưa được tải.');
        }

        this.fileName = fileName;
        this.workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
        
        if (!this.workbook || !this.workbook.SheetNames || this.workbook.SheetNames.length === 0) {
            throw new Error('File Excel không có Sheet hợp lệ.');
        }

        this.selectedSheet = this.workbook.SheetNames[0];
        this.autoDetectStructure();
        this.parseCurrentSheet();
        this.isLoaded = true;

        return this.getWorkbookInfo();
    }

    autoDetectStructure() {
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return;

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        
        // Find best header row (scan first 10 rows)
        let bestHeaderRow = 1;
        let maxNonEmptyCols = 0;
        let bestNameCol = 2;

        for (let r = 0; r <= Math.min(range.e.r, 10); r++) {
            let nonEmpty = 0;
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
                if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
                    nonEmpty++;
                    const txt = String(cell.v).toLowerCase();
                    if (txt.includes('họ và tên') || txt.includes('họ tên') || txt.includes('tên') || txt.includes('họ đệm')) {
                        bestNameCol = c + 1;
                    }
                }
            }
            if (nonEmpty > maxNonEmptyCols) {
                maxNonEmptyCols = nonEmpty;
                bestHeaderRow = r + 1;
            }
        }

        this.headerRow = bestHeaderRow;
        this.nameColumn = bestNameCol;
        this.sampleRow = bestHeaderRow + 1;

        // Load hidden headers from storage if available
        try {
            const saved = localStorage.getItem(`hidden_headers_${this.selectedSheet}`);
            if (saved) {
                this.hiddenHeaders = new Set(JSON.parse(saved));
            }
        } catch (_) {}
    }

    selectSheet(cfg) {
        if (!this.workbook || !this.workbook.Sheets[cfg.SheetName]) {
            throw new Error(`Sheet '${cfg.SheetName}' không tồn tại.`);
        }
        this.selectedSheet = cfg.SheetName;
        if (cfg.HeaderRow) this.headerRow = cfg.HeaderRow;
        if (cfg.NameColumn) this.nameColumn = cfg.NameColumn;
        if (cfg.SampleRow) this.sampleRow = cfg.SampleRow;
        if (cfg.AutoSkipBlank !== undefined) this.autoSkipBlank = cfg.AutoSkipBlank;
        if (cfg.SuggestionsDisabled !== undefined) this.suggestionsDisabled = cfg.SuggestionsDisabled;

        this.parseCurrentSheet();
        return this.getWorkbookInfo();
    }

    parseCurrentSheet() {
        this.records = [];
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return;

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const startRow = this.sampleRow > 0 ? this.sampleRow : this.headerRow; // 0-indexed (starts strictly after sample row)
        let logicalIndex = 0;

        for (let r = startRow; r <= range.e.r; r++) {
            const nameCell = ws[window.XLSX.utils.encode_cell({ r, c: this.nameColumn - 1 })];
            const nameVal = nameCell && nameCell.v !== undefined ? String(nameCell.v).trim() : '';

            // Check if entire row is blank
            let hasAnyData = false;
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
                if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
                    hasAnyData = true;
                    break;
                }
            }

            if (hasAnyData || nameVal) {
                const rowIndex = r + 1;
                const displayName = nameVal || `(Dòng ${rowIndex})`;
                this.records.push({
                    rowIndex,
                    logicalIndex: logicalIndex++,
                    keyDisplay: displayName,
                    tooltipPreview: `Dòng ${rowIndex} - ${displayName}`
                });
            }
        }
    }

    getWorkbookInfo() {
        return {
            isLoaded: this.isLoaded,
            filePath: this.fileName,
            selectedSheet: this.selectedSheet,
            statusMessage: this.isLoaded ? `Đã mở: ${this.fileName} (${this.records.length} bản ghi)` : 'Chưa mở file',
            isMultiSheetMode: false,
            sheets: this.workbook ? this.workbook.SheetNames : [],
            headerRowNumber: this.headerRow,
            nameColumnIndex: this.nameColumn,
            sampleRowThreshold: this.sampleRow,
            isAutoSaveEnabled: this.isAutoSave,
            suggestionsDisabled: this.suggestionsDisabled,
            autoSkipBlankHeaders: this.autoSkipBlank
        };
    }

    getEffectiveCellText(ws, r, c) {
        if (!ws) return '';
        if (ws['!merges']) {
            for (const m of ws['!merges']) {
                if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
                    const anchorCell = ws[window.XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })];
                    return anchorCell && anchorCell.v !== undefined ? String(anchorCell.w !== undefined ? anchorCell.w : anchorCell.v).trim() : '';
                }
            }
        }
        const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
        return cell && cell.v !== undefined ? String(cell.w !== undefined ? cell.w : cell.v).trim() : '';
    }

    getHeaders() {
        if (!this.workbook || !this.selectedSheet) return [];
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return [];

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const headers = [];
        const r = this.headerRow - 1;

        for (let c = range.s.c; c <= range.e.c; c++) {
            const colIdx = c + 1;
            let name = this.getEffectiveCellText(ws, r, c);
            if (this.autoSkipBlank && name.toLowerCase() === 'blank') continue;
            if (!name) name = `Cột ${colIdx}`;

            const isVisible = !this.hiddenHeaders.has(colIdx);
            headers.push({
                columnIndex: colIdx,
                name,
                isVisible,
                sheetName: this.selectedSheet,
                toggleLabel: `Cột ${colIdx}: ${name}`
            });
        }
        return headers;
    }

    saveHiddenHeaders(hiddenIndices) {
        this.hiddenHeaders = new Set(hiddenIndices);
        try {
            localStorage.setItem(`hidden_headers_${this.selectedSheet}`, JSON.stringify(Array.from(this.hiddenHeaders)));
        } catch (_) {}
        return { success: true };
    }

    toggleAllHeaders() {
        this.showAllHeaders = !this.showAllHeaders;
        if (this.showAllHeaders) {
            this.hiddenHeaders.clear();
        }
        return { success: true };
    }

    getRecords(search = '') {
        if (!search) return this.records;
        const q = search.toLowerCase().trim();
        return this.records.filter(r => r.keyDisplay.toLowerCase().includes(q));
    }

    getDataValidationOptionsForCell(sheetName, r, c) {
        if (!this.workbook) return [];
        const targetSheet = sheetName || this.selectedSheet;
        const ws = this.workbook.Sheets[targetSheet];
        if (!ws) return [];

        const validations = ws['!dataValidation'] || ws['!dataValidations'] || [];
        for (const val of validations) {
            if (!val || !val.sqref) continue;
            // Check if (r, c) is covered by sqref range string
            const ranges = String(val.sqref).trim().split(/\s+/);
            let inRange = false;
            for (const rng of ranges) {
                if (!rng) continue;
                try {
                    const dec = window.XLSX.utils.decode_range(rng);
                    if (r >= dec.s.r && r <= dec.e.r && c >= dec.s.c && c <= dec.e.c) {
                        inRange = true;
                        break;
                    }
                } catch (_) {}
            }

            if (!inRange) continue;

            // Only list validations produce dropdown options
            if (val.type === 'list' || !val.type) {
                let f1 = val.formula1 ? String(val.formula1).trim() : '';
                if (!f1) continue;

                // Strip quotes if formula1 is a quoted list like '"Nam,Nữ"' or '"1,2,3"'
                if ((f1.startsWith('"') && f1.endsWith('"')) || (f1.startsWith("'") && f1.endsWith("'"))) {
                    f1 = f1.substring(1, f1.length - 1);
                    return f1.split(/[,;\t]/).map(s => s.trim()).filter(Boolean);
                }

                // If comma separated without quotes: "Nam,Nữ"
                if (!f1.includes('!') && !f1.startsWith('=') && (f1.includes(',') || f1.includes(';'))) {
                    return f1.split(/[,;\t]/).map(s => s.trim()).filter(Boolean);
                }

                // If formula reference e.g. =Sheet2!$A$1:$A$10 or Sheet2!A1:A10 or $A$1:$A$10
                if (f1.startsWith('=')) f1 = f1.substring(1).trim();
                let refSheet = targetSheet;
                let rangeAddress = f1;
                if (f1.includes('!')) {
                    const parts = f1.split('!');
                    refSheet = parts[0].replace(/^['"]|['"]$/g, '');
                    rangeAddress = parts[1];
                }

                const refWs = this.workbook.Sheets[refSheet];
                if (refWs) {
                    try {
                        const cleanRange = rangeAddress.replace(/\$/g, '');
                        const dec = window.XLSX.utils.decode_range(cleanRange);
                        const opts = [];
                        for (let rowIdx = dec.s.r; rowIdx <= dec.e.r; rowIdx++) {
                            for (let colIdx = dec.s.c; colIdx <= dec.e.c; colIdx++) {
                                const targetCell = refWs[window.XLSX.utils.encode_cell({ r: rowIdx, c: colIdx })];
                                if (targetCell && targetCell.v !== undefined) {
                                    const optText = String(targetCell.w !== undefined ? targetCell.w : targetCell.v).trim();
                                    if (optText && !opts.includes(optText)) {
                                        opts.push(optText);
                                    }
                                }
                            }
                        }
                        if (opts.length > 0) return opts;
                    } catch (_) {}
                }
            }
        }
        return [];
    }

    getRecordFields(rowIndex) {
        if (!this.workbook || !this.selectedSheet) return [];
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return [];

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const headers = this.getHeaders();
        const headerMap = new Map(headers.map(h => [h.columnIndex, h]));
        const fields = [];
        const r = rowIndex - 1;
        const shownParentHeaders = new Set();

        // Build distinct suggestions per column
        for (let c = range.s.c; c <= range.e.c; c++) {
            const colIdx = c + 1;
            const headerInfo = headerMap.get(colIdx);
            if (!headerInfo) continue;
            if (!this.showAllHeaders && !headerInfo.isVisible) continue;

            const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
            const val = cell ? (cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '')) : '';

            const effectiveThreshold = this.sampleRow > 0 ? this.sampleRow : (this.headerRow + 1);
            const sampleVal = this.getEffectiveCellText(ws, effectiveThreshold - 1, c);

            // Exclude texts from row 1 up to sampleRowThreshold (header rows + sample row)
            const excludedTexts = new Set();
            for (let rIdx = 0; rIdx < effectiveThreshold; rIdx++) {
                const t = this.getEffectiveCellText(ws, rIdx, c);
                if (t) excludedTexts.add(t.trim().toLowerCase());
            }

            // Suggestions from data rows strictly after sample row
            const distinctVals = [];
            const seen = new Set();
            for (let scanR = effectiveThreshold; scanR <= range.e.r; scanR++) {
                const sVal = this.getEffectiveCellText(ws, scanR, c);
                if (sVal) {
                    const norm = sVal.trim();
                    const lower = norm.toLowerCase();
                    if (!excludedTexts.has(lower) && !seen.has(lower) && norm.length < 100) {
                        seen.add(lower);
                        distinctVals.push(norm);
                    }
                }
            }
            const suggestions = distinctVals.slice(0, 30);

            // Handle multi-line header text (if header explicitly contains \n from multi-tier header)
            let parentHeaderName = '';
            let headerDisplayName = headerInfo.name;
            let showParentHeader = false;

            if (headerInfo.name.includes('\n')) {
                const parts = headerInfo.name.split('\n').map(p => p.trim()).filter(Boolean);
                if (parts.length > 1) {
                    parentHeaderName = parts[0];
                    headerDisplayName = parts.slice(1).join('\n');
                    showParentHeader = !shownParentHeaders.has(parentHeaderName);
                    shownParentHeaders.add(parentHeaderName);
                }
            }

            // Check genuine Excel Data Validation for this cell
            const validationOptions = this.getDataValidationOptionsForCell(this.selectedSheet, r, c);
            const hasValidation = validationOptions.length > 0;
            const suggestionOptions = !hasValidation && !this.suggestionsDisabled ? suggestions : [];

            fields.push({
                headerName: headerInfo.name,
                headerDisplayName: headerDisplayName,
                parentHeaderName,
                showParentHeader,
                isGroupedUnderParentHeader: !!parentHeaderName,
                isFirstInHeaderGroup: false,
                isLastInHeaderGroup: false,
                groupBorderHex: 'Transparent',
                parentTitleBackgroundHex: 'Transparent',
                parentTitleForegroundHex: '#111827',
                columnIndex: colIdx,
                sheetName: this.selectedSheet,
                showSheetSeparator: false,
                sheetSeparatorTitle: '',
                sourceRowIndex: rowIndex,
                value: val,
                sampleValue: sampleVal,
                hasDropdown: hasValidation,
                hasLargeDropdown: hasValidation && validationOptions.length > 10,
                hasSmallDropdown: hasValidation && validationOptions.length <= 10,
                showDropdownEditor: hasValidation,
                isDependentDropdown: false,
                parentDropdownColumns: [],
                dropdownOptions: validationOptions,
                suggestionOptions: suggestionOptions,
                hasSuggestions: suggestionOptions.length > 0,
                isDropdownValueInvalid: false,
                dropdownValidationMessage: '',
                rowHighlightBackgroundHex: 'Transparent',
                rowHighlightBorderHex: 'Transparent',
                rowHeaderForegroundHex: '#111827',
                rowInputBackgroundHex: '#FFFFFF',
                rowInputForegroundHex: '#111827',
                hasRowHighlight: false
            });
        }
        return fields;
    }

    updateCell(sheetName, rowIndex, columnIndex, value) {
        const targetSheet = sheetName || this.selectedSheet;
        const ws = this.workbook.Sheets[targetSheet];
        if (!ws) throw new Error(`Sheet '${targetSheet}' không tồn tại.`);

        const r = rowIndex - 1;
        const c = columnIndex - 1;
        const addr = window.XLSX.utils.encode_cell({ r, c });

        if (!ws[addr]) {
            ws[addr] = { t: 's', v: value };
        } else {
            ws[addr].v = value;
            ws[addr].w = undefined;
            if (typeof value === 'number') {
                ws[addr].t = 'n';
            } else {
                ws[addr].t = 's';
            }
        }

        // Update range
        const range = window.XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        if (r > range.e.r) range.e.r = r;
        if (c > range.e.c) range.e.c = c;
        ws['!ref'] = window.XLSX.utils.encode_range(range);

        // Update record title if name column
        if (columnIndex === this.nameColumn) {
            const rec = this.records.find(item => item.rowIndex === rowIndex);
            if (rec) {
                rec.keyDisplay = value && value.trim() ? value.trim() : `(Dòng ${rowIndex})`;
            }
        }

        return { success: true };
    }

    createRecord(name) {
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws) throw new Error('Chưa mở sheet.');

        const range = window.XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
        const newRowIndex = range.e.r + 2; // 1-indexed

        this.updateCell(this.selectedSheet, newRowIndex, this.nameColumn, name || '');
        
        const newRec = {
            rowIndex: newRowIndex,
            logicalIndex: this.records.length,
            keyDisplay: name || `(Dòng ${newRowIndex})`,
            tooltipPreview: `Dòng ${newRowIndex} - ${name}`
        };
        this.records.push(newRec);

        return newRec;
    }

    deleteRecord(rowIndex) {
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws) throw new Error('Chưa mở sheet.');

        const r = rowIndex - 1;
        const range = window.XLSX.utils.decode_range(ws['!ref']);
        
        // Clear cells in row
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = window.XLSX.utils.encode_cell({ r, c });
            delete ws[addr];
        }

        this.records = this.records.filter(rec => rec.rowIndex !== rowIndex);
        return { success: true };
    }

    renameRecord(rowIndex, logicalIndex, newName) {
        this.updateCell(this.selectedSheet, rowIndex, this.nameColumn, newName);
        return { success: true };
    }

    getSortColumns() {
        const headers = this.getHeaders();
        return headers.map(h => ({
            columnIndex: h.columnIndex,
            name: h.name,
            isCurrentNameColumn: h.columnIndex === this.nameColumn
        }));
    }

    applySort(cfg) {
        const colIdx = cfg.ColumnIndex;
        const mode = cfg.Mode || 0; // 0: Ascending, 1: Descending, 2: Custom
        const ws = this.workbook.Sheets[this.selectedSheet];
        const collator = new Intl.Collator('vi', { numeric: true, sensitivity: 'base' });

        this.records.sort((a, b) => {
            const cellA = ws[window.XLSX.utils.encode_cell({ r: a.rowIndex - 1, c: colIdx - 1 })];
            const cellB = ws[window.XLSX.utils.encode_cell({ r: b.rowIndex - 1, c: colIdx - 1 })];
            const valA = cellA && cellA.v !== undefined ? String(cellA.v) : '';
            const valB = cellB && cellB.v !== undefined ? String(cellB.v) : '';

            let cmp = collator.compare(valA, valB);
            return mode === 1 ? -cmp : cmp;
        });

        // Reassign logicalIndex
        this.records.forEach((r, idx) => r.logicalIndex = idx);
        return { success: true, count: this.records.length };
    }

    downloadWorkbook() {
        if (!this.workbook) throw new Error('Chưa mở workbook.');
        const out = window.XLSX.write(this.workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        return { blob, filename: this.fileName || 'data.xlsx' };
    }
}

const clientEngine = new ClientWorkbookEngine();

// ─── HYBRID REQUEST DISPATCHER ──────────────────────────────────────────────
async function request(endpoint, options = {}) {
    // If client-only mode is active, direct immediately to client engine
    if (isClientOnlyMode) {
        return executeClientFallback(endpoint, options);
    }

    const url = `${API_BASE}${endpoint}`;
    const headers = options.headers || {};

    if (!(options.body instanceof FormData) && options.body && typeof options.body === 'object') {
        headers['Content-Type'] = 'application/json; charset=utf-8';
        options.body = JSON.stringify(options.body);
    }

    options.headers = headers;

    try {
        const response = await fetch(url, options);

        if (response.status === 404) {
            // Server API not present -> Fallback to client engine!
            return executeClientFallback(endpoint, options);
        }

        const contentType = response.headers.get('content-type') || '';
        if (contentType.includes('spreadsheetml') || contentType.includes('octet-stream') || contentType.includes('zip')) {
            if (!response.ok) {
                const text = await response.text();
                throw new Error(text || 'Tải file thất bại');
            }
            const blob = await response.blob();
            return { blob, filename: extractFilename(response) };
        }

        let data = null;
        let errorText = '';
        if (contentType.includes('application/json')) {
            data = await response.json().catch(() => null);
        } else {
            errorText = await response.text().catch(() => '');
            try { data = JSON.parse(errorText); } catch (_) {}
        }

        if (!response.ok) {
            const err = (data && data.error) || (data && data.message) || (typeof data === 'string' ? data : null) || errorText || response.statusText || 'Lỗi xử lý yêu cầu';
            throw new Error(err);
        }

        return data;
    } catch (err) {
        // If network/connection error to backend -> fallback to client engine
        if (err.name === 'TypeError' || err.message.includes('Failed to fetch')) {
            return executeClientFallback(endpoint, options);
        }
        console.error(`API Error on ${endpoint}:`, err);
        throw err;
    }
}

// ─── CLIENT FALLBACK ROUTER ────────────────────────────────────────────────
async function executeClientFallback(endpoint, options = {}) {
    const method = (options.method || 'GET').toUpperCase();
    const body = options.body ? (typeof options.body === 'string' ? JSON.parse(options.body) : options.body) : {};

    // Workbook
    if (endpoint === '/workbook/info') {
        return clientEngine.getWorkbookInfo();
    }
    if (endpoint === '/workbook/open-upload') {
        if (body instanceof FormData) {
            const file = body.get('file');
            return await clientEngine.openFromFile(file);
        }
        throw new Error('Thiếu dữ liệu file.');
    }
    if (endpoint === '/workbook/close') {
        clientEngine.isLoaded = false;
        clientEngine.workbook = null;
        return { message: 'Đã đóng file.' };
    }
    if (endpoint === '/workbook/select-sheet') {
        return clientEngine.selectSheet(body);
    }
    if (endpoint === '/workbook/headers') {
        return clientEngine.getHeaders();
    }
    if (endpoint === '/workbook/headers/hidden') {
        return clientEngine.saveHiddenHeaders(body);
    }
    if (endpoint === '/records/toggle-all-headers') {
        return clientEngine.toggleAllHeaders();
    }
    if (endpoint === '/workbook/save' || endpoint === '/workbook/download') {
        return clientEngine.downloadWorkbook();
    }

    // Records
    if (endpoint.startsWith('/records?') || endpoint === '/records') {
        const urlParams = new URLSearchParams(endpoint.split('?')[1] || '');
        const q = urlParams.get('q') || '';
        return clientEngine.getRecords(q);
    }
    if (endpoint.match(/\/records\/(\d+)\/fields/)) {
        const match = endpoint.match(/\/records\/(\d+)\/fields/);
        const rowIndex = parseInt(match[1], 10);
        return clientEngine.getRecordFields(rowIndex);
    }
    if (endpoint === '/records/cell') {
        return clientEngine.updateCell(body.sheetName, body.rowIndex, body.columnIndex, body.value);
    }
    if (endpoint === '/records' && method === 'POST') {
        return clientEngine.createRecord(body.name);
    }
    if (endpoint.match(/\/records\/(\d+)\/rename/) && method === 'PUT') {
        const match = endpoint.match(/\/records\/(\d+)\/rename/);
        const rowIndex = parseInt(match[1], 10);
        return clientEngine.renameRecord(rowIndex, body.logicalIndex, body.newName);
    }
    if (endpoint.match(/\/records\/(\d+)/) && method === 'DELETE') {
        const match = endpoint.match(/\/records\/(\d+)/);
        const rowIndex = parseInt(match[1], 10);
        return clientEngine.deleteRecord(rowIndex);
    }

    // Settings
    if (endpoint === '/settings/sort-columns') {
        return clientEngine.getSortColumns();
    }
    if (endpoint === '/settings/sort' && method === 'POST') {
        return clientEngine.applySort(body);
    }
    if (endpoint === '/settings/session') {
        return clientEngine.getWorkbookInfo();
    }

    console.warn(`[Client Engine] Endpoint ${endpoint} handled with default response.`);
    return { success: true };
}

function extractFilename(response) {
    const disposition = response.headers.get('content-disposition');
    if (disposition && disposition.includes('filename=')) {
        const match = disposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
        if (match && match[1]) {
            return match[1].replace(/['"]/g, '');
        }
    }
    return 'download.xlsx';
}

export const api = {
    // Workbook
    getWorkbookInfo: () => request('/workbook/info'),
    openPath: (path) => request('/workbook/open-path', { method: 'POST', body: { path } }),
    browseLocal: () => request('/workbook/browse-local', { method: 'POST' }),
    openUpload: (formData) => request('/workbook/open-upload', { method: 'POST', body: formData }),
    closeWorkbook: () => request('/workbook/close', { method: 'POST' }),
    selectSheet: (cfg) => request('/workbook/select-sheet', { method: 'POST', body: cfg }),
    getHeaders: () => request('/workbook/headers'),
    saveHiddenHeaders: (hidden) => request('/workbook/headers/hidden', { method: 'POST', body: hidden }),
    toggleAllHeaders: () => request('/records/toggle-all-headers', { method: 'POST' }),
    saveWorkbook: () => request('/workbook/save', { method: 'POST' }),
    saveWorkbookAs: (newPath) => request('/workbook/save-as', { method: 'POST', body: { path: newPath } }),
    saveAsDialog: () => request('/workbook/save-as-dialog', { method: 'POST' }),
    downloadWorkbook: () => request('/workbook/download'),

    // Records
    getRecords: async (search = '') => {
        const q = search ? `?q=${encodeURIComponent(search)}` : '';
        const items = await request(`/records${q}`);
        return Array.isArray(items) ? items : (items?.items || []);
    },
    getRecordFields: (rowIndex, logicalIndex = -1) => request(`/records/${rowIndex}/fields?logicalIndex=${logicalIndex}`),
    updateCell: (sheetName, rowIndex, columnIndex, value) => request('/records/cell', {
        method: 'POST',
        body: { sheetName: sheetName || '', rowIndex, columnIndex, value: value ?? '' }
    }),
    createRecord: (name) => request('/records', {
        method: 'POST',
        body: { name: name ?? '' }
    }),
    deleteRecord: (rowIndex, logicalIndex = -1) => request(`/records/${rowIndex}?logicalIndex=${logicalIndex}`, {
        method: 'DELETE'
    }),
    renameRecord: (rowIndex, logicalIndex, newName) => request(`/records/${rowIndex}/rename`, {
        method: 'PUT',
        body: { logicalIndex, newName }
    }),

    // Dropdowns
    getDropdown: (columnIndex, rowIndex = 0, sheetName = '') =>
        request(`/dropdown/${columnIndex}?rowIndex=${rowIndex}&sheetName=${encodeURIComponent(sheetName)}`),
    searchDropdown: (columnIndex, q, rowIndex = 0, sheetName = '') =>
        request(`/dropdown/${columnIndex}/search?q=${encodeURIComponent(q)}&rowIndex=${rowIndex}&sheetName=${encodeURIComponent(sheetName)}`),
    getDependentDropdown: (columnIndex, overrides, rowIndex = 0, sheetName = '') => request(`/dropdown/${columnIndex}/dependent`, {
        method: 'POST',
        body: { sheetName, rowIndex, overrides }
    }),

    // Multi-sheet
    getMultiSheetConfig: () => request('/multisheet/config'),
    applyMultiSheet: (session) => request('/multisheet/apply', { method: 'POST', body: session }),
    deleteMultiSheet: () => request('/multisheet', { method: 'DELETE' }),
    getMultiSheetWorksheets: (includeHidden = false) => request(`/multisheet/worksheets?includeHidden=${includeHidden}`),
    getMultiSheetHeaders: (sheet, from, to) => request(`/multisheet/headers?sheet=${encodeURIComponent(sheet)}&from=${from}&to=${to}`),

    // Import List / DataLink
    openImportList: (body) => request('/import-list/open', { method: 'POST', body }),
    getImportListSheets: () => request('/import-list/sheets'),
    getImportListColumns: (filePath, sheetName, headerRow) => {
        const q = new URLSearchParams({ filePath, sheetName, headerRow }).toString();
        return request(`/import-list/columns?${q}`);
    },
    runDataLink: (profile) => request('/import-list/link', { method: 'POST', body: profile }),
    getDataLinkProfiles: () => request('/import-list/profiles'),
    getDataLinkProfile: (name) => request(`/import-list/profiles/${encodeURIComponent(name)}`),
    saveDataLinkProfile: (name, profile) => request(`/import-list/profiles/${encodeURIComponent(name)}`, { method: 'POST', body: profile }),
    deleteDataLinkProfile: (name) => request(`/import-list/profiles/${encodeURIComponent(name)}`, { method: 'DELETE' }),

    // Export HSSK
    getHsskProfile: () => request('/export/hssk/profile'),
    saveHsskProfile: (profile) => request('/export/hssk/profile', { method: 'POST', body: profile }),
    runHsskExport: (req) => request('/export/hssk/run', { method: 'POST', body: req }),

    // Export Import-file
    getImportFileProfile: () => request('/export/import-file/profile'),
    saveImportFileProfile: (profile) => request('/export/import-file/profile', { method: 'POST', body: profile }),
    runImportFileExport: (req) => request('/export/import-file/run', { method: 'POST', body: req }),

    // Settings & Tools
    getSessionSettings: () => request('/settings/session'),
    saveSessionSettings: (settings) => request('/settings/session', { method: 'POST', body: settings }),
    getSortColumns: () => request('/settings/sort-columns'),
    applySort: (cfg) => request('/settings/sort', { method: 'POST', body: cfg }),
    getSkipColumns: () => request('/settings/skip-columns'),
    saveSkipColumns: (columns) => request('/settings/skip-columns', { method: 'POST', body: columns }),
    exportBackup: (outputPath) => request('/config/export-backup', { method: 'POST', body: { outputPath } }),
    importBackup: (formData) => request('/config/import-backup', { method: 'POST', body: formData })
};
