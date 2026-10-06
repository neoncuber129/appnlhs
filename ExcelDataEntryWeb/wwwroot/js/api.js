// API Client for ExcelDataEntryWeb with Full In-Browser Client-Side Engine for GitHub Pages
const API_BASE = '/api';

// Detect if running in pure client-side environment (like GitHub Pages) or local backend
const isClientOnlyMode = window.location.hostname.includes('github.io') || 
                         window.location.protocol === 'file:' || 
                         window.USE_CLIENT_ENGINE === true;

// ─── DATA VALIDATION EXTRACTOR (Reads real Excel List Validations from .xlsx via CFB) ──
function extractXmlFromCFB(cfb, targetPath) {
    if (!window.XLSX || !cfb || !cfb.FileIndex) return '';
    const normalizedTarget = targetPath.replace(/\\/g, '/').replace(/^\//, '').toLowerCase();
    const targetBasename = normalizedTarget.split('/').pop();

    for (const entry of cfb.FileIndex) {
        if (!entry || !entry.name || !entry.content) continue;
        const entryName = entry.name.replace(/\\/g, '/').replace(/^\//, '').toLowerCase();
        if (entryName === normalizedTarget || entryName.endsWith('/' + normalizedTarget)) {
            return typeof entry.content === 'string' ? entry.content : new TextDecoder('utf-8').decode(entry.content);
        }
    }

    for (const entry of cfb.FileIndex) {
        if (!entry || !entry.name || !entry.content) continue;
        const entryName = entry.name.replace(/\\/g, '/').replace(/^\//, '').toLowerCase();
        const entryBasename = entryName.split('/').pop();
        if (entryBasename === targetBasename) {
            return typeof entry.content === 'string' ? entry.content : new TextDecoder('utf-8').decode(entry.content);
        }
    }

    return '';
}

function parseAllDataValidations(buffer, workbook) {
    if (!buffer || !window.XLSX) return {};
    let cfb;
    try {
        const uint8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
        cfb = window.XLSX.CFB.read(uint8, { type: 'array' });
    } catch (e) {
        try {
            cfb = window.XLSX.CFB.read(buffer, { type: 'buffer' });
        } catch (_) {
            return {};
        }
    }
    if (!cfb) return {};

    const relsXml = extractXmlFromCFB(cfb, 'xl/_rels/workbook.xml.rels') || extractXmlFromCFB(cfb, 'workbook.xml.rels');
    const wbXml = extractXmlFromCFB(cfb, 'xl/workbook.xml') || extractXmlFromCFB(cfb, 'workbook.xml');

    const rIdToTarget = {};
    const relRegex = /<Relationship\s+([^>]*?)\/>/gi;
    let rMatch;
    while ((rMatch = relRegex.exec(relsXml)) !== null) {
        const idMatch = rMatch[1].match(/Id="([^"]+)"/i);
        const targetMatch = rMatch[1].match(/Target="([^"]+)"/i);
        if (idMatch && targetMatch) {
            let target = targetMatch[1].replace(/^\//, '');
            if (!target.startsWith('xl/')) target = 'xl/' + target;
            rIdToTarget[idMatch[1]] = target;
        }
    }

    const sheetMap = {};
    const sheetRegex = /<sheet\s+([^>]*?)(?:\/>|>.*?<\/sheet>|>)/gi;
    let sMatch;
    while ((sMatch = sheetRegex.exec(wbXml)) !== null) {
        const nameMatch = sMatch[1].match(/name="([^"]+)"/i);
        const rIdMatch = sMatch[1].match(/(?:r:)?id="([^"]+)"/i);
        if (nameMatch && rIdMatch) {
            const sheetName = nameMatch[1];
            const target = rIdToTarget[rIdMatch[1]];
            if (target) {
                const xml = extractXmlFromCFB(cfb, target);
                if (xml) sheetMap[sheetName] = xml;
            }
        }
    }

    // Fallback if sheetMap is empty
    if (Object.keys(sheetMap).length === 0 && workbook && workbook.SheetNames) {
        workbook.SheetNames.forEach((name, idx) => {
            const xml = extractXmlFromCFB(cfb, `xl/worksheets/sheet${idx + 1}.xml`) || extractXmlFromCFB(cfb, `sheet${idx + 1}.xml`);
            if (xml) sheetMap[name] = xml;
        });
    }

    const validationsBySheet = {};

    for (const [sheetName, sheetXml] of Object.entries(sheetMap)) {
        const dvRules = [];

        // Universal XML matcher for standard, x14, and namespaced DataValidations
        const dvRegex = /<(?:[a-zA-Z0-9_-]+:)?dataValidation\s+([^>]*?)>(.*?)<\/(?:[a-zA-Z0-9_-]+:)?dataValidation>/gsi;
        let match;
        while ((match = dvRegex.exec(sheetXml)) !== null) {
            const attrs = match[1];
            const body = match[2];
            const typeMatch = attrs.match(/type="([^"]+)"/i);
            const sqrefMatch = attrs.match(/sqref="([^"]+)"/i) || body.match(/<(?:[a-zA-Z0-9_-]+:)?sqref>(.*?)<\/(?:[a-zA-Z0-9_-]+:)?sqref>/si);
            const formulaMatch = body.match(/<(?:[a-zA-Z0-9_-]+:)?formula1>(.*?)<\/(?:[a-zA-Z0-9_-]+:)?formula1>/si);

            const isList = !typeMatch || typeMatch[1].toLowerCase() === 'list';
            if (isList && sqrefMatch && formulaMatch) {
                const sqref = (sqrefMatch[1] || '').replace(/<[^>]+>/g, '').trim();
                let formula = formulaMatch[1].replace(/<[^>]+>/g, '').trim();
                const ranges = parseRange(sqref);
                const options = resolveFormulaToValues(formula, workbook, sheetName);
                if (options.length > 0) {
                    dvRules.push({ sqref, formula, ranges, options });
                }
            }
        }

        validationsBySheet[sheetName] = dvRules;
    }

    return validationsBySheet;
}

function parseRange(sqref) {
    if (!sqref || !window.XLSX) return [];
    const tokens = sqref.split(/\s+/).filter(Boolean);
    const result = [];

    for (const token of tokens) {
        try {
            const cleanToken = token.replace(/\$/g, '');
            if (cleanToken.includes(':')) {
                const [start, end] = cleanToken.split(':');
                const s = decodeCellOrCol(start, 0);
                const e = decodeCellOrCol(end, 1048575);
                if (s && e) result.push({ s, e });
            } else if (/^[A-Za-z]+$/.test(cleanToken)) {
                const c = window.XLSX.utils.decode_col(cleanToken);
                result.push({ s: { c, r: 0 }, e: { c, r: 1048575 } });
            } else {
                const cell = window.XLSX.utils.decode_cell(cleanToken);
                if (cell) result.push({ s: cell, e: cell });
            }
        } catch (_) {}
    }
    return result;
}

function decodeCellOrCol(str, defaultRow = 0) {
    str = str.replace(/\$/g, '');
    if (/^[A-Za-z]+$/.test(str)) {
        const c = window.XLSX.utils.decode_col(str);
        return { c, r: defaultRow };
    }
    return window.XLSX.utils.decode_cell(str);
}

function isCellInRanges(ranges, r, c) {
    if (!ranges || ranges.length === 0) return false;
    for (const range of ranges) {
        if (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c) {
            return true;
        }
    }
    return false;
}

function isColumnInRanges(ranges, c) {
    if (!ranges || ranges.length === 0) return false;
    for (const range of ranges) {
        if (c >= range.s.c && c <= range.e.c) {
            return true;
        }
    }
    return false;
}

function resolveFormulaToValues(formula, workbook, currentSheetName) {
    if (!formula) return [];
    formula = formula.trim();

    // Decode XML entities
    formula = formula
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>');

    if (formula.startsWith('=')) formula = formula.substring(1).trim();

    // 1. Quoted string list: "Nam,Nữ" or "Nam;Nữ"
    if (formula.startsWith('"') && formula.endsWith('"')) {
        const inner = formula.substring(1, formula.length - 1);
        const sep = inner.includes(';') ? ';' : ',';
        return inner.split(sep).map(s => s.trim()).filter(Boolean);
    }

    // 2. Unquoted comma/semicolon list: Nam,Nữ (no formulas or cell refs)
    if (!formula.includes('!') && !formula.includes(':') && !formula.includes('(')) {
        const sep = formula.includes(';') ? ';' : (formula.includes(',') ? ',' : null);
        if (sep) {
            return formula.split(sep).map(s => s.trim().replace(/^"|"$/g, '')).filter(Boolean);
        }
    }

    // 3. Named range in workbook.Workbook.Names
    if (workbook && workbook.Workbook && workbook.Workbook.Names) {
        const namedRange = workbook.Workbook.Names.find(n => n.Name && n.Name.toLowerCase() === formula.toLowerCase());
        if (namedRange && namedRange.Ref) {
            return resolveFormulaToValues(namedRange.Ref, workbook, currentSheetName);
        }
    }

    // 4. Cell range: [Sheet!]A1:B10 or Sheet!$A$1:$A$10
    let sheetName = currentSheetName;
    let rangeRef = formula;

    if (formula.includes('!')) {
        const parts = formula.split('!');
        sheetName = parts[0].replace(/^'|'$/g, '');
        rangeRef = parts[1];
    }

    // Handle OFFSET / INDIRECT / complex formula: extract inner range or fallback
    if (rangeRef.includes('(')) {
        const innerMatch = rangeRef.match(/([A-Za-z0-9_]+!\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)/);
        if (innerMatch) {
            return resolveFormulaToValues(innerMatch[1], workbook, currentSheetName);
        }
        const cellMatch = rangeRef.match(/(\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?)/);
        if (cellMatch) {
            rangeRef = cellMatch[1];
        }
    }

    rangeRef = rangeRef.replace(/\$/g, '');

    const targetSheet = workbook ? workbook.Sheets[sheetName] : null;
    if (!targetSheet || !window.XLSX) return [];

    try {
        if (rangeRef.includes(':')) {
            const range = window.XLSX.utils.decode_range(rangeRef);
            const values = [];
            for (let r = range.s.r; r <= range.e.r; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const cell = targetSheet[window.XLSX.utils.encode_cell({ r, c })];
                    if (cell && cell.v !== undefined) {
                        const str = String(cell.v).trim();
                        if (str && !values.includes(str)) {
                            values.push(str);
                        }
                    }
                }
            }
            return values;
        } else {
            const cell = targetSheet[rangeRef];
            if (cell && cell.v !== undefined) {
                const str = String(cell.v).trim();
                return str ? [str] : [];
            }
        }
    } catch (_) {}

    return [];
}

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
        this.validationsBySheet = {};
    }

    async openFromFile(file) {
        if (!file) throw new Error('Không có file để mở.');
        this.fileName = file.name || 'data.xlsx';
        const buffer = await file.arrayBuffer();
        return await this.openFromBuffer(buffer, this.fileName);
    }

    async openFromBuffer(buffer, fileName = 'data.xlsx') {
        if (!window.XLSX) {
            throw new Error('Thư viện SheetJS (XLSX) chưa được tải.');
        }

        this.fileName = fileName;
        this.workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
        
        if (!this.workbook || !this.workbook.SheetNames || this.workbook.SheetNames.length === 0) {
            throw new Error('File Excel không có Sheet hợp lệ.');
        }

        try {
            this.validationsBySheet = parseAllDataValidations(buffer, this.workbook);
            console.log('[DataValidation] Extracted validations:', this.validationsBySheet);
        } catch (e) {
            console.warn('Lỗi đọc data validations:', e);
            this.validationsBySheet = {};
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
        const sheet = cfg.SheetName || cfg.sheetName;
        if (!this.workbook || !this.workbook.Sheets[sheet]) {
            throw new Error(`Sheet '${sheet}' không tồn tại.`);
        }
        this.selectedSheet = sheet;
        if (cfg.HeaderRow || cfg.headerRow) this.headerRow = cfg.HeaderRow || cfg.headerRow;
        if (cfg.NameColumn || cfg.nameColumn) this.nameColumn = cfg.NameColumn || cfg.nameColumn;
        if (cfg.SampleRow || cfg.sampleRow) this.sampleRow = cfg.SampleRow || cfg.sampleRow;
        if (cfg.AutoSkipBlank !== undefined) this.autoSkipBlank = cfg.AutoSkipBlank;
        else if (cfg.autoSkipBlank !== undefined) this.autoSkipBlank = cfg.autoSkipBlank;
        if (cfg.SuggestionsDisabled !== undefined) this.suggestionsDisabled = cfg.SuggestionsDisabled;
        else if (cfg.suggestionsDisabled !== undefined) this.suggestionsDisabled = cfg.suggestionsDisabled;
        if (cfg.AutoSave !== undefined) this.isAutoSave = cfg.AutoSave;
        else if (cfg.autoSave !== undefined) this.isAutoSave = cfg.autoSave;

        this.parseCurrentSheet();
        return this.getWorkbookInfo();
    }

    parseCurrentSheet() {
        this.records = [];
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return;

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const startRow = this.sampleRow - 1; // 0-indexed
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

    getHeaders() {
        if (!this.workbook || !this.selectedSheet) return [];
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return [];

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const headers = [];
        const r = this.headerRow - 1;

        for (let c = range.s.c; c <= range.e.c; c++) {
            const colIdx = c + 1;
            const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
            let name = cell && cell.v !== undefined ? String(cell.v).trim() : '';
            if (this.autoSkipBlank && (!name || name.toLowerCase() === 'blank')) {
                continue;
            }
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

    getRecordFields(rowIndex) {
        if (!this.workbook || !this.selectedSheet) return [];
        const ws = this.workbook.Sheets[this.selectedSheet];
        if (!ws || !ws['!ref']) return [];

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const headers = this.getHeaders();
        const headerMap = new Map(headers.map(h => [h.columnIndex, h]));
        const fields = [];
        const r = rowIndex - 1;

        // Parent header tracking (only if header explicitly contains '\n' from merged/multi-row headers)
        const shownParentHeaders = new Set();

        for (let c = range.s.c; c <= range.e.c; c++) {
            const colIdx = c + 1;
            const headerInfo = headerMap.get(colIdx);
            if (!headerInfo) continue;
            if (!this.showAllHeaders && !headerInfo.isVisible) continue;

            const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
            const val = cell ? (cell.w !== undefined ? String(cell.w) : (cell.v !== undefined ? String(cell.v) : '')) : '';

            // Suggestions from other rows in this column (for autocomplete combobox, NOT forced dropdown validation)
            let suggestions = [];
            if (!this.suggestionsDisabled) {
                const distinctVals = new Set();
                for (let scanR = this.sampleRow - 1; scanR <= range.e.r; scanR++) {
                    const scanCell = ws[window.XLSX.utils.encode_cell({ r: scanR, c })];
                    if (scanCell && scanCell.v !== undefined) {
                        const sVal = String(scanCell.v).trim();
                        if (sVal && sVal.length < 100) distinctVals.add(sVal);
                    }
                }
                suggestions = Array.from(distinctVals).slice(0, 30);
            }

            // Grouping: ONLY if header name contains '\n'
            let headerDisplayName = headerInfo.name;
            let parentHeaderName = '';
            let showParentHeader = false;
            let isGrouped = false;

            if (headerInfo.name.includes('\n')) {
                const parts = headerInfo.name.split('\n').map(p => p.trim()).filter(Boolean);
                if (parts.length > 1) {
                    parentHeaderName = parts[0];
                    headerDisplayName = parts.slice(1).join(' - ');
                    isGrouped = true;
                    if (!shownParentHeaders.has(parentHeaderName)) {
                        shownParentHeaders.add(parentHeaderName);
                        showParentHeader = true;
                    }
                }
            }

            // Check real Data Validation from Excel
            const r0 = rowIndex - 1;
            const c0 = colIdx - 1;
            let excelDropdownOptions = [];
            const sheetValidations = (this.validationsBySheet && this.validationsBySheet[this.selectedSheet]) || [];
            
            for (const rule of sheetValidations) {
                if (isCellInRanges(rule.ranges, r0, c0)) {
                    excelDropdownOptions = rule.options || [];
                    if (excelDropdownOptions.length > 0) break;
                }
            }

            if (excelDropdownOptions.length === 0) {
                for (const rule of sheetValidations) {
                    if (isColumnInRanges(rule.ranges, c0)) {
                        excelDropdownOptions = rule.options || [];
                        if (excelDropdownOptions.length > 0) break;
                    }
                }
            }

            const hasDropdown = excelDropdownOptions.length > 0;

            fields.push({
                headerName: headerInfo.name,
                headerDisplayName: headerDisplayName || headerInfo.name,
                parentHeaderName: parentHeaderName,
                showParentHeader: showParentHeader,
                isGroupedUnderParentHeader: isGrouped,
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
                hasDropdown: hasDropdown,
                hasLargeDropdown: excelDropdownOptions.length > 10,
                hasSmallDropdown: hasDropdown && excelDropdownOptions.length <= 10,
                showDropdownEditor: hasDropdown,
                isDependentDropdown: false,
                parentDropdownColumns: [],
                dropdownOptions: excelDropdownOptions,
                suggestionOptions: suggestions,
                hasSuggestions: !hasDropdown && suggestions.length > 0,
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

    getDropdownOptions(columnIndex, rowIndex = 0, sheetName = '') {
        const targetSheet = sheetName || this.selectedSheet;
        const sheetValidations = (this.validationsBySheet && this.validationsBySheet[targetSheet]) || [];
        const r0 = (rowIndex > 0 ? rowIndex : this.sampleRow) - 1;
        const c0 = columnIndex - 1;
        for (const rule of sheetValidations) {
            if (isCellInRanges(rule.ranges, r0, c0)) {
                return rule.options || [];
            }
        }
        for (const rule of sheetValidations) {
            if (isColumnInRanges(rule.ranges, c0)) {
                return rule.options || [];
            }
        }
        return [];
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
