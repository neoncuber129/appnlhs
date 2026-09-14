// Dynamic Form Component with Windows App Matching UI & Keyboard Navigation
import { api } from '../api.js';
import { debounce } from '../utils/debounce.js';
import { showToast } from '../utils/toast.js';

export class DynamicForm {
    constructor({ onFieldSaved, onHasUnsavedChanges }) {
        this.onFieldSaved = onFieldSaved;
        this.onHasUnsavedChanges = onHasUnsavedChanges;
        this.currentRecord = null;
        this.currentFields = [];

        this.scrollContainer = document.getElementById('fields-scroll-container');
        this.captionEl = document.getElementById('current-entry-caption');

        // Debounced cell updater to session (300ms)
        this.debouncedSave = debounce((rec, field, val) => {
            this.executeSave(rec, field, val);
        }, 300);
    }

    setRecord(record, fields) {
        this.currentRecord = record;
        this.currentFields = fields || [];
        this.renderForm();
    }

    renderForm() {
        if (!this.scrollContainer) return;
        this.scrollContainer.innerHTML = '';

        if (!this.currentRecord || this.currentFields.length === 0) {
            if (this.captionEl) {
                this.captionEl.textContent = 'Chưa chọn dòng — chọn một tên trong danh sách bên trái.';
                this.captionEl.classList.remove('active-entry');
            }
            this.scrollContainer.innerHTML = `
                <div class="empty-fields-state">
                    <p>Vui lòng mở file Excel hoặc chọn một tên bên trái để bắt đầu nhập liệu.</p>
                </div>
            `;
            return;
        }

        if (this.captionEl) {
            const name = this.currentRecord.keyDisplay || `Dòng #${this.currentRecord.rowIndex}`;
            this.captionEl.textContent = `Đang nhập cho "${name}" (Dòng ${this.currentRecord.rowIndex})`;
            this.captionEl.classList.add('active-entry');
        }

        let lastSheet = null;

        this.currentFields.forEach((field, fieldIdx) => {
            // Sheet section divider if in multi-sheet mode or sheet changes
            if (field.showSheetSeparator || (field.sheetName && field.sheetName !== lastSheet && lastSheet !== null)) {
                lastSheet = field.sheetName;
                const sheetBanner = document.createElement('div');
                sheetBanner.className = 'win-sheet-divider';
                sheetBanner.innerHTML = `
                    <span class="sheet-divider-icon">📄</span>
                    <span class="sheet-divider-text">${field.sheetSeparatorTitle || `Sheet: ${field.sheetName}`}</span>
                `;
                this.scrollContainer.appendChild(sheetBanner);
            } else if (!lastSheet && field.sheetName) {
                lastSheet = field.sheetName;
            }

            // Parent Header Bar if field starts a group
            if (field.showParentHeader && field.parentHeaderName) {
                const parentBanner = document.createElement('div');
                parentBanner.className = 'win-parent-header-bar';
                parentBanner.textContent = field.parentHeaderName;
                this.scrollContainer.appendChild(parentBanner);
            }

            // Field Row (Header col | Input col)
            const row = document.createElement('div');
            row.className = 'win-field-row';
            row.dataset.colIndex = field.columnIndex;
            row.dataset.fieldIndex = fieldIdx;

            // Header Column (Left)
            const headerCol = document.createElement('div');
            headerCol.className = 'field-header-col';

            const headerTitle = document.createElement('div');
            headerTitle.className = 'field-header-title';
            headerTitle.textContent = field.headerDisplayName || field.headerName || `Cột ${field.columnIndex}`;
            if (field.isRequired) {
                const req = document.createElement('span');
                req.className = 'field-required-star';
                req.textContent = ' *';
                headerTitle.appendChild(req);
            }
            headerCol.appendChild(headerTitle);

            if (field.sampleValue) {
                const sample = document.createElement('div');
                sample.className = 'field-sample-hint';
                sample.textContent = `Mẫu: ${field.sampleValue}`;
                sample.title = `Giá trị mẫu: ${field.sampleValue}`;
                headerCol.appendChild(sample);
            }

            row.appendChild(headerCol);

            // Input Column (Right)
            const inputCol = document.createElement('div');
            inputCol.className = 'field-input-col';

            const control = this.createControl(field);
            inputCol.appendChild(control);

            row.appendChild(inputCol);

            this.scrollContainer.appendChild(row);
        });

        // Auto-focus the first editable control
        setTimeout(() => {
            const firstInput = this.scrollContainer.querySelector('.win-input-ctrl');
            if (firstInput) {
                firstInput.focus();
                if (typeof firstInput.select === 'function') firstInput.select();
            }
        }, 50);
    }

    createControl(field) {
        const val = field.value ?? '';
        const hasExcelDropdown = Boolean(field.hasDropdown && field.dropdownOptions && field.dropdownOptions.length > 0);

        // 1. Trường có Validation ở file Excel: CHỈ CHO PHÉP CHỌN TRONG CÁC GIÁ TRỊ CÓ SẴN
        if (hasExcelDropdown) {
            const select = document.createElement('select');
            select.className = 'win-input-ctrl win-select-ctrl';
            if (field.hasRowHighlight && field.rowHighlightBorderHex && field.rowHighlightBorderHex !== 'Transparent') {
                select.style.borderColor = field.rowHighlightBorderHex;
            }

            const defaultPrompt = field.sampleValue ? `-- Chọn (Mẫu: ${field.sampleValue}) --` : '-- Chọn giá trị --';
            const emptyOpt = document.createElement('option');
            emptyOpt.value = '';
            emptyOpt.textContent = defaultPrompt;
            select.appendChild(emptyOpt);

            let matched = false;
            const options = field.dropdownOptions || [];
            options.forEach(opt => {
                const optEl = document.createElement('option');
                optEl.value = opt;
                optEl.textContent = opt;
                if (val && (opt.trim().toLowerCase() === val.trim().toLowerCase() || opt === val)) {
                    optEl.selected = true;
                    matched = true;
                }
                select.appendChild(optEl);
            });

            // Nếu giá trị cũ trong file không nằm trong danh sách validate (dữ liệu cũ sai)
            if (val && !matched) {
                const legacyOpt = document.createElement('option');
                legacyOpt.value = val;
                legacyOpt.textContent = `${val} (Hiện tại)`;
                legacyOpt.selected = true;
                select.appendChild(legacyOpt);
            }

            select.addEventListener('change', () => {
                this.handleValueChange(field, select.value, true);
            });

            select.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.handleValueChange(field, select.value, true);
                    this.focusNextField(select);
                }
            });

            return select;
        }

        // 2. Tính năng gợi ý theo cột (không phải validation): cho phép gõ tự do kèm danh sách gợi ý
        if (field.suggestionOptions && field.suggestionOptions.length > 0) {
            const wrapper = document.createElement('div');
            wrapper.className = 'win-combobox-wrapper';

            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'win-input-ctrl win-combobox-input';
            input.value = val;
            input.placeholder = field.sampleValue ? `Mẫu: ${field.sampleValue}` : '';

            // Create unique datalist
            const listId = `dl-${field.columnIndex}-${Math.random().toString(36).substring(2, 7)}`;
            const datalist = document.createElement('datalist');
            datalist.id = listId;

            field.suggestionOptions.forEach(opt => {
                const optEl = document.createElement('option');
                optEl.value = opt;
                datalist.appendChild(optEl);
            });

            input.setAttribute('list', listId);
            wrapper.appendChild(input);
            wrapper.appendChild(datalist);

            // Button to trigger dropdown arrow / open list
            const arrowBtn = document.createElement('button');
            arrowBtn.type = 'button';
            arrowBtn.className = 'combobox-arrow-btn';
            arrowBtn.innerHTML = '▼';
            arrowBtn.tabIndex = -1;
            arrowBtn.addEventListener('click', () => {
                input.focus();
                if (input.value) {
                    input.select();
                }
            });
            wrapper.appendChild(arrowBtn);

            this.attachInputEvents(input, field);
            return wrapper;
        }

        // Multiline textarea
        if (field.isMultiLine) {
            const textarea = document.createElement('textarea');
            textarea.className = 'win-input-ctrl win-textarea';
            textarea.value = val;
            textarea.rows = 2;
            textarea.placeholder = field.sampleValue ? `Mẫu: ${field.sampleValue}` : '';
            this.attachInputEvents(textarea, field);
            return textarea;
        }

        // Regular input
        const input = document.createElement('input');
        input.className = 'win-input-ctrl';
        input.value = val;
        input.placeholder = field.sampleValue ? `Mẫu: ${field.sampleValue}` : '';

        if (field.isNumber) {
            input.type = 'number';
            input.step = 'any';
        } else {
            input.type = 'text';
        }

        this.attachInputEvents(input, field);
        return input;
    }

    attachInputEvents(element, field) {
        // Debounced save on typing
        element.addEventListener('input', () => {
            this.handleValueChange(field, element.value, false);
        });

        // Immediate save on blur
        element.addEventListener('blur', () => {
            this.handleValueChange(field, element.value, true);
        });

        // Enter key navigation -> Moves to NEXT field, exactly like the Windows app!
        element.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                if (element.tagName === 'TEXTAREA' && e.ctrlKey) {
                    // In textarea, Ctrl+Enter keeps newline, Enter moves next
                    return;
                }
                e.preventDefault();
                this.handleValueChange(field, element.value, true);
                this.focusNextField(element);
            }
        });
    }

    focusNextField(currentElement) {
        const allControls = Array.from(this.scrollContainer.querySelectorAll('.win-input-ctrl'));
        const idx = allControls.indexOf(currentElement);
        if (idx >= 0 && idx < allControls.length - 1) {
            const next = allControls[idx + 1];
            next.focus();
            if (typeof next.select === 'function') {
                next.select();
            }
        }
    }

    handleValueChange(field, newValue, immediate = false) {
        if (field.value === newValue) return;
        field.value = newValue;

        // Báo có thay đổi chưa lưu để thanh trạng thái trên cùng hiện cảnh báo
        if (this.onHasUnsavedChanges) {
            this.onHasUnsavedChanges(true);
        }

        const targetRecord = this.currentRecord;
        if (immediate) {
            this.executeSave(targetRecord, field, newValue);
        } else {
            this.debouncedSave(targetRecord, field, newValue);
        }
    }

    async executeSave(record, field, value) {
        const targetRec = record || this.currentRecord;
        if (!targetRec) return;
        const row = field.sourceRowIndex > 0 ? field.sourceRowIndex : targetRec.rowIndex;

        try {
            await api.updateCell(field.sheetName || '', row, field.columnIndex, value);
            if (this.onFieldSaved) {
                this.onFieldSaved(targetRec, field.columnIndex, value);
            }
        } catch (err) {
            showToast(`Lỗi cập nhật ô: ${err.message}`, 'error');
        }
    }
}
