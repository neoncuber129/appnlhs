// Toolbar and Navigation Component (Matching Windows App Top & Sheet Bars)
import { api } from '../api.js';
import { showToast } from '../utils/toast.js';
import { saveActiveFileHandle, getActiveFileHandle, clearActiveFileHandle, isFileLockedError } from '../utils/storage.js';

export class Toolbar {
    constructor({
        onWorkbookOpened,
        onWorkbookClosed,
        onSheetSelect,
        onToggleMultiSheet,
        onExportImportFile,
        onConfigSingleSheet,
        onConfigMultiSheet,
        onSettingsHssk,
        onHeaderSettings,
        onToggleAllHeaders,
        onOpenPathRequest
    }) {
        this.onWorkbookOpened = onWorkbookOpened;
        this.onWorkbookClosed = onWorkbookClosed;
        this.onSheetSelect = onSheetSelect;
        this.onToggleMultiSheet = onToggleMultiSheet;
        this.onExportImportFile = onExportImportFile;
        this.onConfigSingleSheet = onConfigSingleSheet;
        this.onConfigMultiSheet = onConfigMultiSheet;
        this.onSettingsHssk = onSettingsHssk;
        this.onHeaderSettings = onHeaderSettings;
        this.onToggleAllHeaders = onToggleAllHeaders;
        this.onOpenPathRequest = onOpenPathRequest;

        // Top bar buttons
        this.btnBrowse = document.getElementById('btn-browse-file');
        this.btnOpenPath = document.getElementById('btn-open-path');
        this.fallbackFileInput = document.getElementById('fallback-file-input');
        this.btnSaveNow = document.getElementById('btn-save-now');
        this.btnSaveAs = document.getElementById('btn-save-as');
        this.btnDownload = document.getElementById('btn-download-file');
        this.scaleSelect = document.getElementById('scale-select');
        this.btnSettings = document.getElementById('btn-settings-hssk');
        this.activeFilenameEl = document.getElementById('active-filename');
        this.btnCloseWorkbook = document.getElementById('btn-close-workbook');
        this.themeToggle = document.getElementById('btn-theme-toggle');
        this.statusPill = document.getElementById('app-status-pill');

        // Sheet bar controls
        this.sheetDropdown = document.getElementById('sheet-dropdown');
        this.btnToggleMultiSheet = document.getElementById('btn-toggle-multisheet');
        this.multiSheetLabel = document.getElementById('multisheet-btn-label');
        this.btnExportImport = document.getElementById('btn-export-import-file');
        this.btnConfigSingle = document.getElementById('btn-config-single-sheet');
        this.btnConfigMulti = document.getElementById('btn-config-multi-sheet');

        // Right panel entry actions
        this.btnHeaderSettings = document.getElementById('btn-header-settings');
        this.btnToggleAllHeadersEl = document.getElementById('btn-toggle-all-headers');

        this.initEvents();
        this.initTheme();
        this.initScale();
    }

    initEvents() {
        // 1. Chọn file Excel (Mở ngay lập tức hộp thoại chọn file của hệ thống)
        if (this.btnBrowse) {
            this.btnBrowse.addEventListener('click', () => {
                this.handleBrowseFile();
            });
        }

        // 2. Mở theo đường dẫn trực tiếp trên máy
        if (this.btnOpenPath) {
            this.btnOpenPath.addEventListener('click', () => {
                if (this.onOpenPathRequest) {
                    this.onOpenPathRequest();
                }
            });
        }

        if (this.fallbackFileInput) {
            this.fallbackFileInput.addEventListener('change', async (e) => {
                if (e.target.files && e.target.files.length > 0) {
                    window._activeFileHandle = null;
                    await clearActiveFileHandle();
                    await this.uploadFile(e.target.files[0]);
                }
            });
        }

        // Đóng file / Chọn lại từ đầu
        if (this.btnCloseWorkbook) {
            this.btnCloseWorkbook.addEventListener('click', async () => {
                await this.handleCloseWorkbook();
            });
        }

        // Cảnh báo khi người dùng reload / đóng tab nếu có thay đổi chưa lưu
        window.addEventListener('beforeunload', (e) => {
            if (window._hasUnsavedChanges) {
                e.preventDefault();
                e.returnValue = 'Bạn có thay đổi chưa lưu vào file Excel. Bạn có chắc muốn rời đi không?';
                return e.returnValue;
            }
        });

        // 3. Lưu ngay (Ctrl+S & Nút Lưu ngay)
        if (this.btnSaveNow) {
            this.btnSaveNow.addEventListener('click', async () => {
                await this.handleSaveNow();
            });
        }

        window.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
                e.preventDefault();
                if (this.btnSaveNow && !this.btnSaveNow.disabled) {
                    this.handleSaveNow();
                }
            }
        });

        // 4. Save As
        if (this.btnSaveAs) {
            this.btnSaveAs.addEventListener('click', async () => {
                await this.handleSaveAs();
            });
        }

        // 5. Tải bản sao file
        if (this.btnDownload) {
            this.btnDownload.addEventListener('click', async () => {
                await this.handleDownload();
            });
        }

        // 6. Kích thước (Co giãn font chữ)
        if (this.scaleSelect) {
            this.scaleSelect.addEventListener('change', (e) => {
                this.applyScale(e.target.value);
            });
        }

        // 7. Cài đặt HSSK
        if (this.btnSettings) {
            this.btnSettings.addEventListener('click', () => {
                if (this.onSettingsHssk) this.onSettingsHssk();
            });
        }

        // 8. Chuyển sheet
        if (this.sheetDropdown) {
            this.sheetDropdown.addEventListener('change', (e) => {
                const sheet = e.target.value;
                if (sheet && this.onSheetSelect) this.onSheetSelect(sheet);
            });
        }

        // 9. Nhập nhiều sheet
        if (this.btnToggleMultiSheet) {
            this.btnToggleMultiSheet.addEventListener('click', () => {
                if (this.onToggleMultiSheet) this.onToggleMultiSheet();
            });
        }

        // 10. Xuất file import
        if (this.btnExportImport) {
            this.btnExportImport.addEventListener('click', () => {
                if (this.onExportImportFile) this.onExportImportFile();
            });
        }

        // 11. Cấu hình sheet
        if (this.btnConfigSingle) {
            this.btnConfigSingle.addEventListener('click', () => {
                if (this.onConfigSingleSheet) this.onConfigSingleSheet();
            });
        }

        if (this.btnConfigMulti) {
            this.btnConfigMulti.addEventListener('click', () => {
                if (this.onConfigMultiSheet) this.onConfigMultiSheet();
            });
        }

        // 12. Ẩn/Hiện mục nhập liệu & Hiện tất cả
        if (this.btnHeaderSettings) {
            this.btnHeaderSettings.addEventListener('click', () => {
                if (this.onHeaderSettings) this.onHeaderSettings();
            });
        }

        if (this.btnToggleAllHeadersEl) {
            this.btnToggleAllHeadersEl.addEventListener('click', () => {
                if (this.onToggleAllHeaders) this.onToggleAllHeaders();
            });
        }
    }

    async handleBrowseFile() {
        // Ưu tiên 1: Dùng File System Access API trên Chrome/Edge/Cốc Cốc
        // Cho phép mở hộp thoại OS lập tức và có quyền ghi thẳng vào file gốc trên máy
        if (window.showOpenFilePicker) {
            try {
                const [fileHandle] = await window.showOpenFilePicker({
                    types: [{
                        description: 'Excel Workbook (*.xlsx, *.xlsm, *.xlsb, *.xls)',
                        accept: {
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm', '.xlsb', '.xls']
                        }
                    }],
                    multiple: false
                });

                if (!fileHandle) return;

                const file = await fileHandle.getFile();
                window._activeFileHandle = fileHandle; // Lưu file handle vào memory
                await saveActiveFileHandle(fileHandle); // Lưu vào IndexedDB để không mất khi reload!

                await this.uploadFile(file);
                return;
            } catch (err) {
                if (err.name === 'AbortError') {
                    // Người dùng hủy chọn file
                    return;
                }
                console.warn('showOpenFilePicker error, falling back to input:', err);
            }
        }

        // Ưu tiên 2: Sử dụng HTML file input chuẩn
        if (this.fallbackFileInput) {
            this.fallbackFileInput.value = '';
            this.fallbackFileInput.click();
        }
    }

    async uploadFile(file) {
        if (!file) return;
        try {
            this.setStatus(`Đang mở file ${file.name}...`);
            showToast(`Đang mở file ${file.name}...`, 'info');

            const formData = new FormData();
            formData.append('file', file);

            const info = await api.openUpload(formData);
            this.setStatus(`Đã mở file: ${file.name}`);
            showToast(`Đã mở file ${file.name} thành công!`, 'success');
            if (this.onWorkbookOpened) this.onWorkbookOpened(info);
        } catch (err) {
            this.setStatus('Lỗi mở file');
            showToast(`Lỗi mở file: ${err.message}`, 'error');
        }
    }

    async handleSaveNow() {
        try {
            this.setStatus('Đang lưu trực tiếp vào file trên máy...');
            
            // 1. Gọi server lưu dữ liệu
            try {
                await api.saveWorkbook();
            } catch (serverErr) {
                if (isFileLockedError(serverErr)) {
                    this.setStatus('⚠️ File đang mở trong Excel — Hãy đóng Excel!');
                    showToast('⚠️ Không thể lưu: File đang được mở bằng Microsoft Excel hoặc ứng dụng khác!\nVui lòng ĐÓNG FILE trong Excel rồi nhấn "Lưu ngay" lại.', 'error', 8000);
                    return;
                }
                throw serverErr;
            }

            // 2. Lấy file handle (từ memory hoặc khôi phục từ IndexedDB)
            let handle = window._activeFileHandle;
            if (!handle) {
                handle = await getActiveFileHandle();
            }

            if (handle) {
                try {
                    // Kiểm tra và yêu cầu quyền ghi nếu cần (sau khi reload trang)
                    let perm = await handle.queryPermission({ mode: 'readwrite' });
                    if (perm !== 'granted') {
                        perm = await handle.requestPermission({ mode: 'readwrite' });
                    }

                    if (perm === 'granted') {
                        window._activeFileHandle = handle;
                        await saveActiveFileHandle(handle);

                        const { blob } = await api.downloadWorkbook();
                        let writable;
                        try {
                            writable = await handle.createWritable();
                        } catch (writeInitErr) {
                            if (isFileLockedError(writeInitErr)) {
                                this.setStatus('⚠️ File đang mở trong Excel — Hãy đóng Excel!');
                                showToast('⚠️ Không thể lưu: File đang được mở bằng Microsoft Excel hoặc ứng dụng khác!\nVui lòng ĐÓNG FILE trong Excel rồi nhấn "Lưu ngay" lại.', 'error', 8000);
                                return;
                            }
                            throw writeInitErr;
                        }

                        try {
                            await writable.write(blob);
                            await writable.close();
                        } catch (writeCloseErr) {
                            if (isFileLockedError(writeCloseErr)) {
                                this.setStatus('⚠️ File đang mở trong Excel — Hãy đóng Excel!');
                                showToast('⚠️ Không thể lưu: File đang được mở bằng Microsoft Excel hoặc ứng dụng khác!\nVui lòng ĐÓNG FILE trong Excel rồi nhấn "Lưu ngay" lại.', 'error', 8000);
                                return;
                            }
                            throw writeCloseErr;
                        }

                        this.notifyUnsavedChanges(false);
                        const saveTime = new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
                        this.setStatus(`Đã lưu trực tiếp vào file máy tính (${saveTime})`);
                        showToast('Đã lưu trực tiếp vào file Excel trên máy tính!', 'success');
                        return;
                    } else {
                        this.setStatus('Chưa cấp quyền ghi');
                        showToast('Chưa được cấp quyền ghi vào file gốc trên máy. Vui lòng cấp quyền để lưu file!', 'warning');
                        return;
                    }
                } catch (handleErr) {
                    console.warn('Ghi trực tiếp file handle:', handleErr);
                    if (isFileLockedError(handleErr)) {
                        this.setStatus('⚠️ File đang mở trong Excel — Hãy đóng Excel!');
                        showToast('⚠️ Không thể lưu: File đang được mở bằng Microsoft Excel hoặc ứng dụng khác!\nVui lòng ĐÓNG FILE trong Excel rồi nhấn "Lưu ngay" lại.', 'error', 8000);
                        return;
                    }
                    showToast(`Lỗi ghi vào file máy tính: ${handleErr.message || handleErr.name}`, 'warning');
                    return;
                }
            }

            // 3. Nếu không có file handle (mất liên kết hoàn toàn):
            // Mở hộp thoại Save As để chọn nơi lưu file trên máy tính
            if (window.showSaveFilePicker) {
                showToast('Chưa có liên kết với file trên máy. Vui lòng chọn nơi lưu file!', 'info');
                await this.handleSaveAs();
            } else {
                await this.handleDownload();
                this.notifyUnsavedChanges(false);
                this.setStatus('Đã tải bản sao file về máy');
                showToast('Đã tải bản sao file về máy!', 'success');
            }
        } catch (err) {
            if (isFileLockedError(err)) {
                this.setStatus('⚠️ File đang mở trong Excel — Hãy đóng Excel!');
                showToast('⚠️ Không thể lưu: File đang được mở bằng Microsoft Excel hoặc ứng dụng khác!\nVui lòng ĐÓNG FILE trong Excel rồi nhấn "Lưu ngay" lại.', 'error', 8000);
            } else {
                this.setStatus('Lỗi lưu file');
                showToast(`Lỗi lưu file: ${err.message}`, 'error');
            }
        }
    }

    async handleSaveAs() {
        // Cho phép chọn nơi lưu file mới trực tiếp qua showSaveFilePicker
        if (window.showSaveFilePicker) {
            try {
                const suggested = this.activeFilenameEl?.textContent || 'data_entry.xlsx';
                const handle = await window.showSaveFilePicker({
                    suggestedName: suggested,
                    types: [{
                        description: 'Excel Workbook (*.xlsx)',
                        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                    }]
                });

                if (!handle) return;

                this.setStatus('Đang lưu bản sao file mới...');
                const { blob } = await api.downloadWorkbook();
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();

                // Chuyển quyền active handle sang file mới
                window._activeFileHandle = handle;
                await saveActiveFileHandle(handle);

                const newFile = await handle.getFile();
                if (this.activeFilenameEl) {
                    this.activeFilenameEl.textContent = newFile.name;
                    this.activeFilenameEl.title = newFile.name;
                }
                this.setStatus(`Đã lưu thành file: ${newFile.name}`);
                showToast(`Đã lưu thành file mới: ${newFile.name}`, 'success');
                return;
            } catch (err) {
                if (err.name === 'AbortError') return;
                console.warn('showSaveFilePicker error:', err);
                if (isFileLockedError(err)) {
                    this.setStatus('⚠️ File đang mở trong Excel — Hãy đóng Excel!');
                    showToast('⚠️ Không thể lưu vào file này: File đang được mở bởi ứng dụng khác (như Excel). Hãy đóng file đó rồi thử lại!', 'error', 7000);
                    return;
                }
                showToast(`Lỗi ghi file mới: ${err.message}`, 'error');
                return;
            }
        }

        // Fallback tải về
        await this.handleDownload();
    }

    async handleCloseWorkbook() {
        try {
            this.setStatus('Đang đóng file...');
            await api.closeWorkbook();
            window._activeFileHandle = null;
            await clearActiveFileHandle();
            this.setWorkbookState(null);
            this.setStatus('Sẵn sàng. Nhấn "Chọn file Excel" để mở file.');
            showToast('Đã đóng file. Bạn có thể chọn file mới!', 'info');
            if (this.onWorkbookClosed) this.onWorkbookClosed();
        } catch (err) {
            console.warn('Lỗi đóng file:', err);
        }
    }

    async handleDownload() {
        try {
            this.setStatus('Đang tải bản sao file về máy...');
            const { blob, filename } = await api.downloadWorkbook();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename || 'data_entry.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
            this.setStatus('Tải file hoàn tất');
            showToast('Đã tải bản sao file Excel về máy!', 'success');
        } catch (err) {
            showToast(`Lỗi tải file: ${err.message}`, 'error');
        }
    }

    initTheme() {
        const saved = localStorage.getItem('app-theme') || 'dark';
        document.documentElement.setAttribute('data-theme', saved);
        this.updateThemeIcon(saved);

        if (this.themeToggle) {
            this.themeToggle.addEventListener('click', () => {
                const cur = document.documentElement.getAttribute('data-theme') || 'dark';
                const next = cur === 'dark' ? 'light' : 'dark';
                document.documentElement.setAttribute('data-theme', next);
                localStorage.setItem('app-theme', next);
                this.updateThemeIcon(next);
            });
        }
    }

    updateThemeIcon(theme) {
        if (!this.themeToggle) return;
        this.themeToggle.innerHTML = theme === 'dark'
            ? `<svg viewBox="0 0 20 20" fill="currentColor" width="16" height="16"><path fill-rule="evenodd" d="M10 2a1 1 0 011 1v1a1 1 0 11-2 0V3a1 1 0 011-1zm4 8a4 4 0 11-8 0 4 4 0 018 0zm-.464 4.95l.707.707a1 1 0 001.414-1.414l-.707-.707a1 1 0 00-1.414 1.414zm2.12-10.607a1 1 0 010 1.414l-.706.707a1 1 0 11-1.414-1.414l.707-.707a1 1 0 011.414 0zM17 11a1 1 0 100-2h-1a1 1 0 100 2h1zm-7 4a1 1 0 011 1v1a1 1 0 11-2 0v-1a1 1 0 011-1zM5.05 6.464A1 1 0 106.465 5.05l-.708-.707a1 1 0 00-1.414 1.414l.707.707zm1.414 8.486l-.707.707a1 1 0 01-1.414-1.414l.707-.707a1 1 0 011.414 1.414zM4 11a1 1 0 100-2H3a1 1 0 000 2h1z" clip-rule="evenodd" /></svg>`
            : `<svg viewBox="0 0 20 20" fill="currentColor" width="16" height="16"><path d="M17.293 13.293A8 8 0 016.707 2.707a8.001 8.001 0 1010.586 10.586z" /></svg>`;
    }

    initScale() {
        const saved = localStorage.getItem('app-scale') || 'medium';
        if (this.scaleSelect) {
            this.scaleSelect.value = saved;
        }
        this.applyScale(saved);
    }

    applyScale(scale) {
        document.documentElement.setAttribute('data-scale', scale);
        localStorage.setItem('app-scale', scale);
    }

    setWorkbookState(info) {
        const isLoaded = info && info.isLoaded;

        if (this.btnSaveNow) this.btnSaveNow.disabled = !isLoaded;
        if (this.btnSaveAs) this.btnSaveAs.disabled = !isLoaded;
        if (this.btnDownload) this.btnDownload.disabled = !isLoaded;
        if (this.sheetDropdown) this.sheetDropdown.disabled = !isLoaded;
        if (this.btnToggleMultiSheet) this.btnToggleMultiSheet.disabled = !isLoaded;
        if (this.btnExportImport) this.btnExportImport.disabled = !isLoaded;
        if (this.btnCloseWorkbook) this.btnCloseWorkbook.style.display = isLoaded ? 'inline-block' : 'none';

        if (this.activeFilenameEl) {
            if (isLoaded && info.filePath) {
                const name = info.filePath.split(/[\\/]/).pop();
                this.activeFilenameEl.textContent = name;
                this.activeFilenameEl.title = info.filePath;
            } else {
                this.activeFilenameEl.textContent = 'Chưa mở file';
                this.activeFilenameEl.title = '';
            }
        }

        if (this.sheetDropdown && isLoaded && info.sheets) {
            this.sheetDropdown.innerHTML = '';
            info.sheets.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === info.selectedSheet) opt.selected = true;
                this.sheetDropdown.appendChild(opt);
            });
        }

        if (this.multiSheetLabel) {
            this.multiSheetLabel.textContent = info && info.isMultiSheetMode
                ? 'Nhập nhiều sheet · Bật'
                : 'Nhập nhiều sheet · Tắt';
        }

        if (this.btnToggleMultiSheet) {
            if (info && info.isMultiSheetMode) {
                this.btnToggleMultiSheet.classList.add('active');
            } else {
                this.btnToggleMultiSheet.classList.remove('active');
            }
        }

        if (this.btnConfigMulti) {
            this.btnConfigMulti.style.display = (info && info.isMultiSheetMode) ? 'inline-block' : 'none';
        }

        this.notifyUnsavedChanges(false);
    }

    setStatus(text) {
        if (this.statusPill) {
            this.statusPill.textContent = text;
        }
    }

    notifyUnsavedChanges(hasUnsaved) {
        window._hasUnsavedChanges = !!hasUnsaved;
        if (hasUnsaved) {
            if (this.statusPill) {
                this.statusPill.textContent = '● Có thay đổi chưa lưu — Nhấn "Lưu ngay" (Ctrl+S)';
                this.statusPill.classList.add('unsaved');
                this.statusPill.title = 'Bạn vừa sửa dữ liệu. Hãy nhấn Lưu ngay hoặc Ctrl+S để lưu vào file máy tính.';
            }
            if (this.btnSaveNow) {
                this.btnSaveNow.classList.add('has-unsaved');
                this.btnSaveNow.title = 'Có dữ liệu mới chưa lưu! Nhấn để lưu ngay (Ctrl+S)';
            }
        } else {
            if (this.statusPill) {
                this.statusPill.classList.remove('unsaved');
                this.statusPill.title = '';
            }
            if (this.btnSaveNow) {
                this.btnSaveNow.classList.remove('has-unsaved');
                this.btnSaveNow.title = 'Lưu trực tiếp vào file Excel trên máy (Ctrl+S)';
            }
        }
    }
}
