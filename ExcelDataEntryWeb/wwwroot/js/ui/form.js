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
        const dropdownOpts = (field.dropdownOptions && field.dropdownOptions.length > 0) ? field.dropdownOptions : [];
        const hasExcelDropdown = field.hasDropdown && dropdownOpts.length > 0;

        // 1. CHỈ TRƯỜNG THỰC SỰ CÓ DROPDOWN VALIDATE TRONG EXCEL MỚI HIỂN THỊ GIAO DIỆN DROPDOWN/COMBOBOX
        if (hasExcelDropdown) {
            return this.createComboboxControl(field, dropdownOpts, val);
        }

        // 2. Multiline textarea
        if (field.isMultiLine) {
            const textarea = document.createElement('textarea');
            textarea.className = 'win-input-ctrl win-textarea';
            textarea.value = val;
            textarea.rows = 2;
            textarea.placeholder = field.sampleValue ? `Mẫu: ${field.sampleValue}` : '';
            this.attachInputEvents(textarea, field);
            return textarea;
        }

        // 3. Regular input
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

    createComboboxControl(field, options, val) {
        const wrapper = document.createElement('div');
        wrapper.className = 'win-combobox-wrapper';
        if (field.hasRowHighlight && field.rowHighlightBorderHex && field.rowHighlightBorderHex !== 'Transparent') {
            wrapper.style.borderColor = field.rowHighlightBorderHex;
        }

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'win-input-ctrl win-combobox-input';
        input.value = val;
        input.placeholder = field.sampleValue ? `Mẫu: ${field.sampleValue}` : 'Chọn hoặc nhập giá trị...';
        input.autocomplete = 'off';

        const arrowBtn = document.createElement('button');
        arrowBtn.type = 'button';
        arrowBtn.className = 'combobox-arrow-btn';
        arrowBtn.innerHTML = '▼';
        arrowBtn.title = 'Bấm để mở danh sách chọn';
        arrowBtn.tabIndex = -1;

        let popover = null;
        let highlightedIndex = -1;

        const closePopover = () => {
            if (popover) {
                popover.remove();
                popover = null;
                highlightedIndex = -1;
            }
        };

        const renderPopoverItems = (filteredList) => {
            if (!popover) return;
            popover.innerHTML = '';

            if (filteredList.length === 0) {
                popover.innerHTML = '<div class="win-dropdown-empty">Không tìm thấy mục phù hợp</div>';
                return;
            }

            filteredList.forEach((opt, idx) => {
                const itemEl = document.createElement('div');
                itemEl.className = 'win-dropdown-option-item';
                if (opt.trim().toLowerCase() === (input.value || '').trim().toLowerCase()) {
                    itemEl.classList.add('selected');
                }
                if (idx === highlightedIndex) {
                    itemEl.classList.add('highlighted');
                }
                itemEl.textContent = opt;

                itemEl.addEventListener('mousedown', (e) => {
                    e.preventDefault(); // Prevent input blur before click
                });

                itemEl.addEventListener('click', (e) => {
                    e.stopPropagation();
                    input.value = opt;
                    this.handleValueChange(field, opt, true);
                    closePopover();
                    input.focus();
                });

                popover.appendChild(itemEl);
            });
        };

        const openPopover = () => {
            closePopover();
            popover = document.createElement('div');
            popover.className = 'win-dropdown-popover';

            const q = (input.value || '').trim().toLowerCase();
            const filtered = q ? options.filter(o => o.toLowerCase().includes(q)) : options;
            renderPopoverItems(filtered.length > 0 ? filtered : options);

            wrapper.appendChild(popover);
        };

        const togglePopover = (e) => {
            if (e) e.stopPropagation();
            if (popover) {
                closePopover();
            } else {
                openPopover();
                input.focus();
            }
        };

        arrowBtn.addEventListener('click', togglePopover);

        input.addEventListener('click', () => {
            if (!popover) {
                openPopover();
            }
        });

        input.addEventListener('input', () => {
            this.handleValueChange(field, input.value, false);
            if (!popover) {
                openPopover();
            } else {
                const q = input.value.trim().toLowerCase();
                const filtered = q ? options.filter(o => o.toLowerCase().includes(q)) : options;
                renderPopoverItems(filtered.length > 0 ? filtered : options);
            }
        });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown') {
                if (!popover) {
                    e.preventDefault();
                    openPopover();
                    return;
                }
                e.preventDefault();
                const items = popover.querySelectorAll('.win-dropdown-option-item');
                if (items.length > 0) {
                    highlightedIndex = Math.min(highlightedIndex + 1, items.length - 1);
                    items.forEach((it, i) => it.classList.toggle('highlighted', i === highlightedIndex));
                    items[highlightedIndex]?.scrollIntoView({ block: 'nearest' });
                }
            } else if (e.key === 'ArrowUp') {
                if (popover) {
                    e.preventDefault();
                    const items = popover.querySelectorAll('.win-dropdown-option-item');
                    if (items.length > 0) {
                        highlightedIndex = Math.max(highlightedIndex - 1, 0);
                        items.forEach((it, i) => it.classList.toggle('highlighted', i === highlightedIndex));
                        items[highlightedIndex]?.scrollIntoView({ block: 'nearest' });
                    }
                }
            } else if (e.key === 'Enter') {
                if (popover && highlightedIndex >= 0) {
                    e.preventDefault();
                    const items = popover.querySelectorAll('.win-dropdown-option-item');
                    if (items[highlightedIndex]) {
                        items[highlightedIndex].click();
                        return;
                    }
                }
                closePopover();
                this.handleValueChange(field, input.value, true);
                this.focusNextField(input);
            } else if (e.key === 'Escape') {
                if (popover) {
                    e.preventDefault();
                    closePopover();
                }
            } else if (e.key === 'Tab') {
                closePopover();
            }
        });

        input.addEventListener('blur', () => {
            setTimeout(() => {
                closePopover();
                this.handleValueChange(field, input.value, true);
            }, 180);
        });

        wrapper.appendChild(input);
        wrapper.appendChild(arrowBtn);
        return wrapper;
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
