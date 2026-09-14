// API Client for ExcelDataEntryWeb
const API_BASE = '/api';

async function request(endpoint, options = {}) {
    const url = `${API_BASE}${endpoint}`;
    const headers = options.headers || {};

    if (!(options.body instanceof FormData) && options.body && typeof options.body === 'object') {
        headers['Content-Type'] = 'application/json; charset=utf-8';
        options.body = JSON.stringify(options.body);
    }

    options.headers = headers;

    try {
        const response = await fetch(url, options);

        // Check if file download
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
        console.error(`API Error on ${endpoint}:`, err);
        throw err;
    }
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
