// Modals Manager for all dialog workflows
import { api } from '../api.js';
import { showToast } from '../utils/toast.js';

export class ModalsManager {
    constructor({ onWorkbookOpened, onRecordsChanged }) {
        this.onWorkbookOpened = onWorkbookOpened;
        this.onRecordsChanged = onRecordsChanged;

        this.modalContainer = document.getElementById('modal-container');
        this.modalTitle = document.getElementById('modal-title');
        this.modalBody = document.getElementById('modal-body');
        this.modalFooter = document.getElementById('modal-footer');
        this.modalCloseBtn = document.getElementById('modal-close-btn');

        this.initEvents();
    }

    initEvents() {
        if (this.modalCloseBtn) {
            this.modalCloseBtn.addEventListener('click', () => this.close());
        }

        // Close when clicking modal backdrop
        if (this.modalContainer) {
            this.modalContainer.addEventListener('click', (e) => {
                if (e.target === this.modalContainer) {
                    this.close();
                }
            });
        }

        window.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this.modalContainer?.classList.contains('active')) {
                this.close();
            }
        });
    }

    open(title, renderContentFn, renderFooterFn) {
        this.modalTitle.textContent = title;
        this.modalBody.innerHTML = '';
        this.modalFooter.innerHTML = '';

        if (renderContentFn) renderContentFn(this.modalBody);
        if (renderFooterFn) renderFooterFn(this.modalFooter);

        this.modalContainer.classList.add('active');
    }

    close() {
        if (this.modalContainer) {
            this.modalContainer.classList.remove('active');
        }
    }



    // ─── MODAL 2: THÊM BẢN GHI MỚI ─────────────────────────────────────────
    showAddRecordModal(currentRecordIndex) {
        this.open('Thêm Bản Ghi Mới', (body) => {
            body.innerHTML = `
                <div class="form-group">
                    <label class="form-label">Tên bản ghi / Người mới:</label>
                    <input type="text" id="add-record-name" class="form-input" placeholder="Nhập họ và tên..." autofocus>
                </div>
                <div class="form-group" style="margin-top: 1rem;">
                    <label class="form-checkbox-label">
                        <input type="checkbox" id="add-insert-after" checked>
                        <span>Chèn ngay sau bản ghi hiện tại (#${currentRecordIndex + 1})</span>
                    </label>
                </div>
            `;
        }, (footer) => {
            const btn = document.createElement('button');
            btn.className = 'btn btn-primary';
            btn.textContent = 'Thêm bản ghi';

            btn.addEventListener('click', async () => {
                const name = document.getElementById('add-record-name').value.trim();
                const insertAfter = document.getElementById('add-insert-after').checked;
                if (!name) {
                    showToast('Vui lòng nhập tên bản ghi!', 'warning');
                    return;
                }

                try {
                    btn.disabled = true;
                    await api.createRecord(name, insertAfter);
                    this.close();
                    showToast(`Đã thêm bản ghi: ${name}`, 'success');
                    if (this.onRecordsChanged) this.onRecordsChanged();
                } catch (err) {
                    showToast(`Lỗi thêm bản ghi: ${err.message}`, 'error');
                    btn.disabled = false;
                }
            });

            footer.appendChild(btn);
        });
    }

    // ─── MODAL 3: EXPORT HSSK ────────────────────────────────────────────
    showHsskExportModal() {
        this.open('Xuất File HSSK', async (body) => {
            body.innerHTML = '<div class="loading-spinner">Đang tải cấu hình HSSK...</div>';
            try {
                const [profile, headers] = await Promise.all([
                    api.getHsskProfile(),
                    api.getHeaders()
                ]);

                let genderOptions = '<option value="-1">-- Không dùng cột giới tính --</option>';
                headers.forEach(h => {
                    const sel = h.columnIndex === profile.genderColumnIndex ? 'selected' : '';
                    genderOptions += `<option value="${h.columnIndex}" ${sel}>Cột ${h.columnIndex}: ${h.name}</option>`;
                });

                body.innerHTML = `
                    <div class="form-group">
                        <label class="form-label">Cột giới tính:</label>
                        <select id="hssk-gender-col" class="form-input">${genderOptions}</select>
                    </div>
                    <div class="form-group" style="margin-top: 1rem;">
                        <label class="form-label">Chọn các cột KHÔNG áp dụng dòng mẫu cho nữ:</label>
                        <div class="checkbox-scroll-list" id="hssk-skip-cols"></div>
                    </div>
                `;

                const skipList = body.querySelector('#hssk-skip-cols');
                const skipSet = new Set(profile.skipSampleColumnIndexes || []);
                headers.forEach(h => {
                    const label = document.createElement('label');
                    label.className = 'form-checkbox-label';
                    label.innerHTML = `
                        <input type="checkbox" value="${h.columnIndex}" ${skipSet.has(h.columnIndex) ? 'checked' : ''}>
                        <span>Cột ${h.columnIndex}: ${h.name}</span>
                    `;
                    skipList.appendChild(label);
                });
            } catch (err) {
                body.innerHTML = `<div class="error-box">Lỗi: ${err.message}</div>`;
            }
        }, (footer) => {
            const runBtn = document.createElement('button');
            runBtn.className = 'btn btn-primary';
            runBtn.textContent = 'Bắt đầu Xuất HSSK';

            runBtn.addEventListener('click', async () => {
                try {
                    runBtn.disabled = true;
                    runBtn.textContent = 'Đang xử lý xuất...';

                    const genderCol = parseInt(document.getElementById('hssk-gender-col').value, 10);
                    const skipBoxes = document.querySelectorAll('#hssk-skip-cols input[type="checkbox"]:checked');
                    const skipCols = Array.from(skipBoxes).map(b => parseInt(b.value, 10));

                    const { blob, filename } = await api.runHsskExport({
                        genderColumnIndex: genderCol,
                        skipSampleColumnIndexes: skipCols
                    });

                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename || 'HSSK_export.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    a.remove();

                    this.close();
                    showToast('Đã xuất file HSSK thành công!', 'success');
                } catch (err) {
                    showToast(`Lỗi xuất HSSK: ${err.message}`, 'error');
                    runBtn.disabled = false;
                    runBtn.textContent = 'Bắt đầu Xuất HSSK';
                }
            });

            footer.appendChild(runBtn);
        });
    }

    // ─── MODAL 4: EXPORT IMPORT FILE ─────────────────────────────────────
    showImportFileExportModal() {
        this.open('Xuất Báo Cáo Import File', async (body) => {
            body.innerHTML = '<div class="loading-spinner">Đang chuẩn bị...</div>';
            try {
                const [profile, headers] = await Promise.all([
                    api.getImportFileProfile(),
                    api.getHeaders()
                ]);

                let genderOptions = '<option value="-1">-- Không dùng cột giới tính --</option>';
                headers.forEach(h => {
                    const sel = h.columnIndex === profile.genderColumnIndex ? 'selected' : '';
                    genderOptions += `<option value="${h.columnIndex}" ${sel}>Cột ${h.columnIndex}: ${h.name}</option>`;
                });

                body.innerHTML = `
                    <div class="form-group">
                        <label class="form-label">Cột giới tính:</label>
                        <select id="imp-gender-col" class="form-input">${genderOptions}</select>
                    </div>
                `;
            } catch (err) {
                body.innerHTML = `<div class="error-box">Lỗi: ${err.message}</div>`;
            }
        }, (footer) => {
            const runBtn = document.createElement('button');
            runBtn.className = 'btn btn-primary';
            runBtn.textContent = 'Xuất File Ngay';

            runBtn.addEventListener('click', async () => {
                try {
                    runBtn.disabled = true;
                    runBtn.textContent = 'Đang xuất...';
                    const genderCol = parseInt(document.getElementById('imp-gender-col').value, 10);

                    const { blob, filename } = await api.runImportFileExport({
                        genderColumnIndex: genderCol
                    });

                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename || 'ImportFile_export.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    a.remove();

                    this.close();
                    showToast('Đã xuất file thành công!', 'success');
                } catch (err) {
                    showToast(`Lỗi xuất: ${err.message}`, 'error');
                    runBtn.disabled = false;
                }
            });

            footer.appendChild(runBtn);
        });
    }

    // ─── MODAL 5: SẮP XẾP BẢN GHI (SORT) ──────────────────────────────────
    showSortModal() {
        this.open('Sắp Xếp Danh Sách Bản Ghi', async (body) => {
            body.innerHTML = '<div class="loading-spinner">Đang tải cột...</div>';
            try {
                const sortColumns = await api.getSortColumns();
                let colOptions = '';
                sortColumns.forEach(c => {
                    colOptions += `<option value="${c.columnIndex}">Cột ${c.columnIndex}: ${c.displayName}</option>`;
                });

                body.innerHTML = `
                    <div class="form-group">
                        <label class="form-label">Chọn cột để sắp xếp:</label>
                        <select id="sort-col-select" class="form-input">${colOptions}</select>
                    </div>
                    <div class="form-group" style="margin-top: 1rem;">
                        <label class="form-label">Chế độ sắp xếp:</label>
                        <select id="sort-mode-select" class="form-input">
                            <option value="Standard">Tiêu chuẩn (Chữ cái A-Z, Số tăng dần)</option>
                            <option value="Natural">Tự nhiên (1, 2, 10 thay vì 1, 10, 2)</option>
                            <option value="Custom">Tùy biến thứ tự</option>
                        </select>
                    </div>
                `;
            } catch (err) {
                body.innerHTML = `<div class="error-box">Lỗi: ${err.message}</div>`;
            }
        }, (footer) => {
            const btn = document.createElement('button');
            btn.className = 'btn btn-primary';
            btn.textContent = 'Thực hiện sắp xếp';

            btn.addEventListener('click', async () => {
                try {
                    btn.disabled = true;
                    const colIndex = parseInt(document.getElementById('sort-col-select').value, 10);
                    const mode = document.getElementById('sort-mode-select').value;

                    await api.applySort({ columnIndex: colIndex, mode });
                    this.close();
                    showToast('Đã sắp xếp danh sách bản ghi!', 'success');
                    if (this.onRecordsChanged) this.onRecordsChanged();
                } catch (err) {
                    showToast(`Lỗi sắp xếp: ${err.message}`, 'error');
                    btn.disabled = false;
                }
            });

            footer.appendChild(btn);
        });
    }

    // ─── MODAL 6: SAO LƯU VÀ CẤU HÌNH ─────────────────────────────────────
    showBackupModal() {
        this.open('Sao Lưu & Phục Hồi Cấu Hình', (body) => {
            body.innerHTML = `
                <div class="backup-section">
                    <h4>1. Xuất file sao lưu (Export Backup)</h4>
                    <p class="section-desc">Sao lưu toàn bộ profile, cấu hình ghép sheet và danh sách bỏ qua mẫu.</p>
                    <div class="form-group">
                        <label class="form-label">Đường dẫn file xuất ra trên máy:</label>
                        <input type="text" id="backup-out-path" class="form-input" placeholder="C:\\Users\\AD\\Desktop\\backup.json">
                    </div>
                    <button id="btn-do-export-backup" class="btn btn-secondary" style="margin-top: 0.5rem;">Xuất bản sao lưu</button>
                </div>

                <hr style="margin: 1.5rem 0; border: 0; border-top: 1px solid var(--border-color);">

                <div class="backup-section">
                    <h4>2. Phục hồi từ file (Import Backup)</h4>
                    <p class="section-desc">Khôi phục cấu hình đã lưu trước đó.</p>
                    <input type="file" id="backup-in-file" class="form-input" accept=".json">
                    <button id="btn-do-import-backup" class="btn btn-secondary" style="margin-top: 0.5rem;">Khôi phục</button>
                </div>
            `;

            body.querySelector('#btn-do-export-backup').addEventListener('click', async () => {
                const path = document.getElementById('backup-out-path').value.trim();
                if (!path) {
                    showToast('Vui lòng nhập đường dẫn file xuất!', 'warning');
                    return;
                }
                try {
                    await api.exportBackup(path);
                    showToast('Đã xuất bản sao lưu thành công!', 'success');
                } catch (err) {
                    showToast(`Lỗi sao lưu: ${err.message}`, 'error');
                }
            });

            body.querySelector('#btn-do-import-backup').addEventListener('click', async () => {
                const fileInput = document.getElementById('backup-in-file');
                if (!fileInput.files || fileInput.files.length === 0) {
                    showToast('Vui lòng chọn file sao lưu .json!', 'warning');
                    return;
                }
                try {
                    const fd = new FormData();
                    fd.append('file', fileInput.files[0]);
                    await api.importBackup(fd);
                    showToast('Khôi phục cấu hình thành công!', 'success');
                } catch (err) {
                    showToast(`Lỗi khôi phục: ${err.message}`, 'error');
                }
            });
        }, (footer) => {
            const closeBtn = document.createElement('button');
            closeBtn.className = 'btn btn-secondary';
            closeBtn.textContent = 'Đóng';
            closeBtn.addEventListener('click', () => this.close());
            footer.appendChild(closeBtn);
        });
    }

    // ─── MODAL 7: SỬA TÊN BẢN GHI (RENAME) ──────────────────────────────────
    showRenameRecordModal(record) {
        if (!record) {
            showToast('Chưa chọn bản ghi nào để sửa tên!', 'warning');
            return;
        }

        const oldName = record.keyDisplay || '';
        this.open('Sửa Tên Bản Ghi Đã Chọn', (body) => {
            body.innerHTML = `
                <div class="form-group">
                    <label class="form-label">Tên bản ghi hiện tại (Dòng ${record.rowIndex}):</label>
                    <input type="text" id="rename-input" class="win-input" value="${oldName}" autofocus style="width:100%; margin-top: 0.5rem; padding: 0.5rem;">
                </div>
            `;
            setTimeout(() => {
                const inp = body.querySelector('#rename-input');
                if (inp) {
                    inp.focus();
                    inp.select();
                }
            }, 50);
        }, (footer) => {
            const btn = document.createElement('button');
            btn.className = 'btn btn-primary';
            btn.textContent = 'Lưu tên mới';

            btn.addEventListener('click', async () => {
                const newName = document.getElementById('rename-input').value.trim();
                if (!newName) {
                    showToast('Vui lòng nhập tên mới!', 'warning');
                    return;
                }
                try {
                    btn.disabled = true;
                    await api.renameRecord(record.rowIndex, record.logicalIndex ?? -1, newName);
                    this.close();
                    showToast(`Đã đổi tên thành: ${newName}`, 'success');
                    if (this.onRecordsChanged) this.onRecordsChanged();
                } catch (err) {
                    showToast(`Lỗi sửa tên: ${err.message}`, 'error');
                    btn.disabled = false;
                }
            });

            footer.appendChild(btn);
        });
    }

    // ─── MODAL 8: ẨN / HIỆN MỤC NHẬP LIỆU (HEADER SETTINGS) ─────────────────
    showHeaderSettingsModal(onSaved) {
        this.open('Ẩn / Hiện Mục Nhập Liệu', async (body) => {
            body.innerHTML = '<div class="loading-spinner">Đang tải danh sách cột...</div>';
            try {
                const headers = await api.getHeaders();
                body.innerHTML = `
                    <div style="margin-bottom: 0.75rem; font-size: 0.875rem; color: var(--text-muted);">
                        Chọn các mục muốn hiển thị trên form nhập liệu:
                    </div>
                    <div class="checkbox-scroll-list" id="headers-check-list" style="max-height: 320px; overflow-y: auto; border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 0.5rem;"></div>
                `;

                const listEl = body.querySelector('#headers-check-list');
                headers.forEach(h => {
                    const row = document.createElement('label');
                    row.className = 'form-checkbox-label';
                    row.style.display = 'flex';
                    row.style.alignItems = 'center';
                    row.style.gap = '0.5rem';
                    row.style.padding = '0.25rem 0.5rem';
                    row.innerHTML = `
                        <input type="checkbox" value="${h.columnIndex}" ${h.isVisible ? 'checked' : ''}>
                        <span><strong>Cột ${h.columnIndex}:</strong> ${h.name}</span>
                    `;
                    listEl.appendChild(row);
                });
            } catch (err) {
                body.innerHTML = `<div class="error-box">Lỗi: ${err.message}</div>`;
            }
        }, (footer) => {
            const btn = document.createElement('button');
            btn.className = 'btn btn-primary';
            btn.textContent = 'Lưu cấu hình hiển thị';

            btn.addEventListener('click', async () => {
                try {
                    btn.disabled = true;
                    const unchecked = document.querySelectorAll('#headers-check-list input[type="checkbox"]:not(:checked)');
                    const hiddenIndexes = Array.from(unchecked).map(box => parseInt(box.value, 10));

                    await api.saveHiddenHeaders(hiddenIndexes);
                    this.close();
                    showToast('Đã lưu cấu hình ẩn/hiện mục nhập liệu!', 'success');
                    if (onSaved) onSaved();
                } catch (err) {
                    showToast(`Lỗi lưu hiển thị: ${err.message}`, 'error');
                    btn.disabled = false;
                }
            });

            footer.appendChild(btn);
        });
    }

    // ─── MODAL 9: LIÊN KẾT SỐ LIỆU ─────────────────────────────────────────
    showDataLinkModal() {
        this.open('Liên Kết Số Liệu (Data Link)', (body) => {
            body.innerHTML = `
                <div style="font-size: 0.9rem; line-height: 1.6;">
                    <p>Chức năng <strong>Liên kết số liệu</strong> cho phép đối chiếu và nhập tự động số liệu từ file Excel ngoài dựa trên các cột khóa (Họ tên, Năm sinh, v.v.).</p>
                    <div style="margin-top: 1rem;">
                        <button id="btn-open-datalink-profile" class="btn btn-secondary">Quản lý cấu hình liên kết</button>
                    </div>
                </div>
            `;
        }, (footer) => {
            const closeBtn = document.createElement('button');
            closeBtn.className = 'btn btn-secondary';
            closeBtn.textContent = 'Đóng';
            closeBtn.addEventListener('click', () => this.close());
            footer.appendChild(closeBtn);
        });
    }

    // ─── MODAL 10: IMPORT DANH SÁCH ───────────────────────────────────────
    showImportListModal() {
        this.open('Import Danh Sách Từ File Khác', (body) => {
            body.innerHTML = `
                <div style="font-size: 0.9rem; line-height: 1.6;">
                    <p>Chức năng <strong>Import danh sách</strong> cho phép nạp thêm danh sách tên từ một file Excel khác vào danh sách hiện tại.</p>
                    <div class="form-group" style="margin-top: 1rem;">
                        <label class="form-label">Chọn file danh sách Excel:</label>
                        <input type="file" id="import-list-file" class="win-input" accept=".xlsx,.xls">
                    </div>
                </div>
            `;
        }, (footer) => {
            const closeBtn = document.createElement('button');
            closeBtn.className = 'btn btn-secondary';
            closeBtn.textContent = 'Đóng';
            closeBtn.addEventListener('click', () => this.close());
            footer.appendChild(closeBtn);
        });
    }

    // ─── MODAL 11: MỞ FILE THEO ĐƯỜNG DẪN TRỰC TIẾP ────────────────────────
    showOpenPathModal(onOpened) {
        this.open('Mở File Trực Tiếp Theo Đường Dẫn Máy Tính', (body) => {
            body.innerHTML = `
                <div style="font-size: 0.9rem; line-height: 1.5; margin-bottom: 1rem;">
                    Nhập hoặc dán đường dẫn file Excel trên máy tính của bạn:
                </div>
                <div class="form-group">
                    <label class="form-label">Đường dẫn file (.xlsx, .xlsm, .xls):</label>
                    <input type="text" id="open-file-path-input" class="win-input" placeholder="C:\\Users\\...\\file.xlsx" style="width: 100%; margin-top: 0.4rem;" autofocus>
                </div>
                <div style="margin-top: 1rem; font-size: 0.85rem; color: var(--text-muted);">
                    <em>Mẹo: Bạn có thể click chuột phải vào file trong Windows Explorer và chọn <strong>Copy as path</strong> để dán vào đây.</em>
                </div>
            `;
            setTimeout(() => {
                const inp = body.querySelector('#open-file-path-input');
                if (inp) inp.focus();
            }, 50);
        }, (footer) => {
            const btn = document.createElement('button');
            btn.className = 'btn btn-primary';
            btn.textContent = 'Mở file';

            btn.addEventListener('click', async () => {
                const path = document.getElementById('open-file-path-input').value.trim();
                if (!path) {
                    showToast('Vui lòng nhập đường dẫn file!', 'warning');
                    return;
                }
                try {
                    btn.disabled = true;
                    btn.textContent = 'Đang mở...';
                    const info = await api.openPath(path);
                    window._activeFileHandle = null; // Direct disk path handled by server!
                    this.close();
                    showToast(`Đã mở file: ${info.filePath.split(/[\\\\/]/).pop()}`, 'success');
                    if (onOpened) onOpened(info);
                } catch (err) {
                    showToast(`Lỗi mở file: ${err.message}`, 'error');
                    btn.disabled = false;
                    btn.textContent = 'Mở file';
                }
            });

            footer.appendChild(btn);
        });
    }

    // ─── MODAL 12: CẤU HÌNH NHẬP TỪNG SHEET (Giống bản Windows 100%) ───────
    showSingleSheetConfigModal(currentInfo, onSaved) {
        const headerRow = currentInfo?.headerRowNumber || 3;
        const nameCol = currentInfo?.nameColumnIndex || 2;
        const sampleRow = currentInfo?.sampleRowThreshold || 4;
        const autoSkip = currentInfo?.autoSkipBlankHeaders ?? true;
        const enableSuggestions = !(currentInfo?.suggestionsDisabled ?? false);
        const autoSave = currentInfo?.isAutoSaveEnabled ?? true;

        this.open('Cấu hình nhập từng sheet', (body) => {
            body.innerHTML = `
                <div style="font-size: 0.875rem; color: var(--text-muted); margin-bottom: 1.25rem;">
                    Cấu hình riêng cho chế độ nhập từng sheet. Không ảnh hưởng cấu hình nhập nhiều sheet.
                </div>
                <div class="form-group" style="margin-bottom: 0.85rem;">
                    <label class="form-label" style="font-weight: 600; margin-bottom: 0.35rem; display: block;">Hàng tiêu đề là hàng số:</label>
                    <input type="number" id="cfg-header-row" class="win-input" value="${headerRow}" min="1" style="width: 100%;">
                </div>
                <div class="form-group" style="margin-bottom: 0.85rem;">
                    <label class="form-label" style="font-weight: 600; margin-bottom: 0.35rem; display: block;">Cột tên (mặc định 2):</label>
                    <input type="number" id="cfg-name-col" class="win-input" value="${nameCol}" min="1" style="width: 100%;">
                </div>
                <div class="form-group" style="margin-bottom: 1.1rem;">
                    <label class="form-label" style="font-weight: 600; margin-bottom: 0.35rem; display: block;">Dòng mẫu là dòng số:</label>
                    <input type="number" id="cfg-sample-row" class="win-input" value="${sampleRow}" min="0" style="width: 100%;">
                </div>
                <div class="form-group" style="display: flex; flex-direction: column; gap: 0.6rem;">
                    <label class="form-checkbox-label" style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                        <input type="checkbox" id="cfg-auto-skip" ${autoSkip ? 'checked' : ''}>
                        <span>Tự bỏ qua header = blank</span>
                    </label>
                    <label class="form-checkbox-label" style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                        <input type="checkbox" id="cfg-suggestions" ${enableSuggestions ? 'checked' : ''}>
                        <span>Bật gợi ý theo cột (có thể lag nếu dữ liệu lớn)</span>
                    </label>
                    <label class="form-checkbox-label" style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                        <input type="checkbox" id="cfg-auto-save" ${autoSave ? 'checked' : ''}>
                        <span>Tự động lưu khi sửa</span>
                    </label>
                </div>
            `;
        }, (footer) => {
            const cancelBtn = document.createElement('button');
            cancelBtn.className = 'btn btn-secondary';
            cancelBtn.textContent = 'Hủy';
            cancelBtn.addEventListener('click', () => this.close());

            const okBtn = document.createElement('button');
            okBtn.className = 'btn btn-primary';
            okBtn.textContent = 'OK';

            okBtn.addEventListener('click', async () => {
                const hRow = parseInt(document.getElementById('cfg-header-row').value, 10) || 3;
                const nCol = parseInt(document.getElementById('cfg-name-col').value, 10) || 2;
                const sRow = parseInt(document.getElementById('cfg-sample-row').value, 10) || 4;
                const skip = document.getElementById('cfg-auto-skip').checked;
                const sugg = !document.getElementById('cfg-suggestions').checked; // suggestionsDisabled
                const save = document.getElementById('cfg-auto-save').checked;

                try {
                    okBtn.disabled = true;
                    okBtn.textContent = 'Đang lưu...';
                    const sheet = currentInfo?.selectedSheet || '';
                    const info = await api.selectSheet({
                        sheetName: sheet,
                        headerRow: hRow,
                        nameColumn: nCol,
                        sampleRow: sRow,
                        autoSkipBlank: skip,
                        suggestionsDisabled: sugg,
                        autoSave: save
                    });

                    this.close();
                    showToast('Đã áp dụng cấu hình sheet mới!', 'success');
                    if (onSaved) onSaved(info);
                } catch (err) {
                    showToast(`Lỗi áp dụng cấu hình: ${err.message}`, 'error');
                    okBtn.disabled = false;
                    okBtn.textContent = 'OK';
                }
            });

            footer.appendChild(cancelBtn);
            footer.appendChild(okBtn);
        });
    }
}


