// Main Application Controller (Matching Windows App Behavior & Workflow)
import { api } from './api.js';
import { showToast } from './utils/toast.js';
import { Toolbar } from './ui/toolbar.js';
import { Sidebar } from './ui/sidebar.js';
import { DynamicForm } from './ui/form.js';
import { ModalsManager } from './ui/modals.js';
import { getActiveFileHandle, clearActiveFileHandle } from './utils/storage.js';

class App {
    constructor() {
        this.selectedRecord = null;
        this.records = [];
        this.workbookInfo = null;

        this.init();
    }

    async init() {
        // 1. Dynamic Form
        this.form = new DynamicForm({
            onFieldSaved: (record, colIdx, val) => {
                const nameCol = this.workbookInfo?.nameColumnIndex || 2;
                if (colIdx === nameCol) {
                    const newTitle = val && val.trim() ? val.trim() : `(Dòng ${record.rowIndex})`;
                    record.keyDisplay = newTitle;
                    this.sidebar.updateRecordTitle(record.rowIndex, newTitle);
                    if (this.selectedRecord && this.selectedRecord.rowIndex === record.rowIndex) {
                        this.selectedRecord.keyDisplay = newTitle;
                    }
                    if (this.form.captionEl) {
                        this.form.captionEl.textContent = `Đang nhập cho "${newTitle}" (Dòng ${record.rowIndex})`;
                    }
                }
            },
            onHasUnsavedChanges: (hasUnsaved) => this.toolbar?.notifyUnsavedChanges(hasUnsaved)
        });

        // 2. Sidebar with 6 actions
        this.sidebar = new Sidebar({
            onRecordSelected: (record) => this.handleRecordSelected(record),
            onRecordAdd: () => this.modals.showAddRecordModal(this.selectedRecord ? this.selectedRecord.rowIndex : 0),
            onRecordEdit: (record) => this.modals.showRenameRecordModal(record || this.selectedRecord),
            onRecordDelete: (record) => this.handleRecordDelete(record || this.selectedRecord),
            onSortClick: () => this.modals.showSortModal(),
            onDataLinkClick: () => this.modals.showDataLinkModal(),
            onImportListClick: () => this.modals.showImportListModal()
        });

        // 3. Modals Manager
        this.modals = new ModalsManager({
            onRecordsChanged: () => {
                this.toolbar?.notifyUnsavedChanges(true);
                this.loadRecords();
            }
        });

        // 4. Toolbar & Sheet bar
        this.toolbar = new Toolbar({
            onWorkbookOpened: (info) => this.handleWorkbookOpened(info),
            onWorkbookClosed: () => this.handleWorkbookClosed(),
            onSheetSelect: (sheetName) => this.switchSheet(sheetName),
            onToggleMultiSheet: () => this.handleToggleMultiSheet(),
            onExportImportFile: () => this.modals.showImportFileExportModal(),
            onConfigSingleSheet: () => this.modals.showSingleSheetConfigModal(this.workbookInfo, (info) => this.handleWorkbookOpened(info)),
            onConfigMultiSheet: () => showToast('Cấu hình nhiều sheet', 'info'),
            onSettingsHssk: () => this.modals.showHsskExportModal(),
            onHeaderSettings: () => this.modals.showHeaderSettingsModal(() => this.reloadCurrentRecordFields()),
            onToggleAllHeaders: () => this.handleToggleAllHeaders(),
            onOpenPathRequest: () => this.modals.showOpenPathModal((info) => this.handleWorkbookOpened(info))
        });

        // 5. Global Drag & Drop handling (Fallback if dragging a file into the browser window)
        this.initDragAndDrop();

        // 6. Check existing session on load
        await this.checkInitialState();
    }

    initDragAndDrop() {
        ['dragenter', 'dragover'].forEach(name => {
            window.addEventListener(name, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        ['dragleave', 'drop'].forEach(name => {
            window.addEventListener(name, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        window.addEventListener('drop', async (e) => {
            const dt = e.dataTransfer;
            if (dt && dt.files && dt.files.length > 0) {
                const file = dt.files[0];
                await this.toolbar.uploadFile(file);
            }
        });
    }

    async checkInitialState() {
        try {
            // Khôi phục file handle từ IndexedDB nếu có
            const storedHandle = await getActiveFileHandle();
            if (storedHandle) {
                window._activeFileHandle = storedHandle;
            }

            const info = await api.getWorkbookInfo();
            if (info && info.isLoaded) {
                // Nếu reload mà KHÔNG có file handle (mất liên kết với file gốc trên máy tính):
                // Reset session và để người dùng chọn file lại từ đầu để đảm bảo lưu được 100%
                if (!window._activeFileHandle) {
                    console.log('Không có liên kết file máy tính sau reload -> Reset chọn file từ đầu.');
                    await api.closeWorkbook();
                    await clearActiveFileHandle();
                    this.handleWorkbookClosed();
                    this.toolbar.setWorkbookState(null);
                    this.toolbar.setStatus('Sẵn sàng. Nhấn "Chọn file Excel" để mở file.');
                    return;
                }

                this.handleWorkbookOpened(info);
                this.toolbar.setStatus(`Đã mở file: ${info.filePath ? info.filePath.split(/[\\/]/).pop() : 'Excel'}`);
            } else {
                await clearActiveFileHandle();
                this.handleWorkbookClosed();
                this.toolbar.setWorkbookState(null);
                this.toolbar.setStatus('Sẵn sàng. Nhấn "Chọn file Excel" để mở file.');
            }
        } catch (err) {
            console.warn('Initial session check:', err);
            this.toolbar.setStatus('Sẵn sàng');
        }
    }

    handleWorkbookClosed() {
        this.workbookInfo = null;
        this.selectedRecord = null;
        this.records = [];
        this.sidebar.setRecords([]);
        this.form.setRecord(null, []);
    }

    async handleWorkbookOpened(info) {
        this.workbookInfo = info;
        this.toolbar.setWorkbookState(info);
        await this.loadRecords();
    }

    async switchSheet(sheetName) {
        try {
            this.toolbar.setStatus(`Đang chuyển sang sheet "${sheetName}"...`);
            const info = await api.selectSheet({
                sheetName,
                headerRow: this.workbookInfo?.headerRowNumber || 3,
                nameColumn: this.workbookInfo?.nameColumnIndex || 2,
                sampleRow: this.workbookInfo?.sampleRowThreshold || 4,
                autoSkipBlank: this.workbookInfo?.autoSkipBlankHeaders ?? true,
                suggestionsDisabled: this.workbookInfo?.suggestionsDisabled ?? false,
                autoSave: this.workbookInfo?.isAutoSaveEnabled ?? true
            });
            this.workbookInfo = info;
            this.toolbar.setWorkbookState(info);
            await this.loadRecords();
            showToast(`Đã chuyển sang sheet "${sheetName}"`, 'info');
        } catch (err) {
            showToast(`Lỗi chuyển sheet: ${err.message}`, 'error');
        }
    }

    async handleToggleMultiSheet() {
        if (!this.workbookInfo) return;
        const currentMode = this.workbookInfo.isMultiSheetMode;
        if (currentMode) {
            // Turn off
            try {
                await api.deleteMultiSheet();
                this.workbookInfo.isMultiSheetMode = false;
                this.toolbar.setWorkbookState(this.workbookInfo);
                await this.loadRecords();
                showToast('Đã tắt chế độ Nhập nhiều sheet', 'info');
            } catch (err) {
                showToast(`Lỗi tắt nhiều sheet: ${err.message}`, 'error');
            }
        } else {
            showToast('Đang bật chế độ Nhập nhiều sheet...', 'info');
            // Can open config or toggle
        }
    }

    async handleToggleAllHeaders() {
        try {
            const res = await api.toggleAllHeaders();
            const labelBtn = document.getElementById('btn-toggle-all-headers');
            if (labelBtn && res && res.label) {
                labelBtn.textContent = res.label;
            }
            await this.reloadCurrentRecordFields();
            showToast('Đã cập nhật hiển thị tất cả các mục', 'info');
        } catch (err) {
            showToast(`Lỗi: ${err.message}`, 'error');
        }
    }

    async loadRecords() {
        try {
            this.records = await api.getRecords('');
            this.sidebar.setRecords(this.records);
            this.toolbar.setStatus(`Đang hiển thị ${this.records.length} tên trong sheet "${this.workbookInfo?.selectedSheet || ''}"`);

            // Auto select first record like Windows app
            if (this.records.length > 0) {
                this.sidebar.selectRecord(this.records[0]);
            } else {
                this.form.setRecord(null, []);
            }
        } catch (err) {
            showToast(`Lỗi tải danh sách tên: ${err.message}`, 'error');
        }
    }

    async handleRecordSelected(record) {
        this.selectedRecord = record;
        await this.reloadCurrentRecordFields();
    }

    async reloadCurrentRecordFields() {
        if (!this.selectedRecord) return;
        try {
            const fields = await api.getRecordFields(this.selectedRecord.rowIndex, this.selectedRecord.logicalIndex ?? -1);
            this.form.setRecord(this.selectedRecord, fields);
        } catch (err) {
            showToast(`Lỗi tải ô nhập liệu: ${err.message}`, 'error');
        }
    }

    async handleRecordDelete(record) {
        if (!record) return;
        const name = record.keyDisplay || `Dòng #${record.rowIndex}`;
        if (!confirm(`Bạn có chắc chắn muốn xóa dòng "${name}" không?`)) {
            return;
        }

        try {
            await api.deleteRecord(record.rowIndex, record.logicalIndex ?? -1);
            this.toolbar?.notifyUnsavedChanges(true);
            showToast(`Đã xóa dòng "${name}"`, 'success');
            await this.loadRecords();
        } catch (err) {
            showToast(`Lỗi xóa dòng: ${err.message}`, 'error');
        }
    }
}

// Start application when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.__app = new App();
});
