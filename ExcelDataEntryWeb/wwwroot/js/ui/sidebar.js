// Sidebar Records List Component (Matching Windows App Left Panel)
export class Sidebar {
    constructor({
        onRecordSelected,
        onRecordAdd,
        onRecordEdit,
        onRecordDelete,
        onSortClick,
        onDataLinkClick,
        onImportListClick
    }) {
        this.onRecordSelected = onRecordSelected;
        this.onRecordAdd = onRecordAdd;
        this.onRecordEdit = onRecordEdit;
        this.onRecordDelete = onRecordDelete;
        this.onSortClick = onSortClick;
        this.onDataLinkClick = onDataLinkClick;
        this.onImportListClick = onImportListClick;

        this.records = [];
        this.filteredRecords = [];
        this.selectedRecord = null;
        this.searchQuery = '';

        this.listboxEl = document.getElementById('records-listbox');
        this.searchInput = document.getElementById('quick-search-input');

        // 6 Action buttons
        this.btnAdd = document.getElementById('btn-add-name');
        this.btnEdit = document.getElementById('btn-edit-name');
        this.btnDelete = document.getElementById('btn-delete-name');
        this.btnSort = document.getElementById('btn-sort-columns');
        this.btnLink = document.getElementById('btn-link-data');
        this.btnImport = document.getElementById('btn-import-list');

        this.initEvents();
    }

    initEvents() {
        if (this.searchInput) {
            this.searchInput.addEventListener('input', (e) => {
                this.searchQuery = e.target.value.trim().toLowerCase();
                this.filterAndRender();
            });
        }

        if (this.btnAdd) {
            this.btnAdd.addEventListener('click', () => {
                if (this.onRecordAdd) this.onRecordAdd();
            });
        }

        if (this.btnEdit) {
            this.btnEdit.addEventListener('click', () => {
                if (this.selectedRecord && this.onRecordEdit) {
                    this.onRecordEdit(this.selectedRecord);
                }
            });
        }

        if (this.btnDelete) {
            this.btnDelete.addEventListener('click', () => {
                if (this.selectedRecord && this.onRecordDelete) {
                    this.onRecordDelete(this.selectedRecord);
                }
            });
        }

        if (this.btnSort) {
            this.btnSort.addEventListener('click', () => {
                if (this.onSortClick) this.onSortClick();
            });
        }

        if (this.btnLink) {
            this.btnLink.addEventListener('click', () => {
                if (this.onDataLinkClick) this.onDataLinkClick();
            });
        }

        if (this.btnImport) {
            this.btnImport.addEventListener('click', () => {
                if (this.onImportListClick) this.onImportListClick();
            });
        }

        // Global Arrow keys navigation between records
        window.addEventListener('keydown', (e) => {
            if (['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement?.tagName)) return;
            if (e.key === 'ArrowUp') {
                e.preventDefault();
                this.navigateRelative(-1);
            } else if (e.key === 'ArrowDown') {
                e.preventDefault();
                this.navigateRelative(1);
            }
        });
    }

    setRecords(records, keepSelected = true) {
        this.records = records || [];
        this.filterAndRender();

        if (keepSelected && this.selectedRecord) {
            const found = this.records.find(r => r.rowIndex === this.selectedRecord.rowIndex);
            if (found) {
                this.selectRecord(found, false);
                return;
            }
        }

        if (this.records.length > 0) {
            this.selectRecord(this.records[0], true);
        } else {
            this.selectedRecord = null;
        }
    }

    filterAndRender() {
        if (!this.searchQuery) {
            this.filteredRecords = [...this.records];
        } else {
            this.filteredRecords = this.records.filter(r =>
                (r.keyDisplay && r.keyDisplay.toLowerCase().includes(this.searchQuery)) ||
                (r.tooltipPreview && r.tooltipPreview.toLowerCase().includes(this.searchQuery))
            );
        }
        this.renderList();
    }

    renderList() {
        if (!this.listboxEl) return;
        this.listboxEl.innerHTML = '';

        if (this.filteredRecords.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'listbox-empty';
            empty.textContent = this.records.length === 0 ? 'Chưa có bản ghi nào' : 'Không tìm thấy tên nào phù hợp';
            this.listboxEl.appendChild(empty);
            return;
        }

        this.filteredRecords.forEach((record, idx) => {
            const isSelected = this.selectedRecord && this.selectedRecord.rowIndex === record.rowIndex;
            const item = document.createElement('div');
            item.className = 'win-listbox-item' + (isSelected ? ' selected' : '');
            item.dataset.rowIndex = record.rowIndex;
            item.tabIndex = 0;

            const orderNum = document.createElement('span');
            orderNum.className = 'item-order-num';
            orderNum.textContent = `${idx + 1}.`;

            const nameSpan = document.createElement('span');
            nameSpan.className = 'item-name-text';
            nameSpan.textContent = record.keyDisplay || `Dòng #${record.rowIndex}`;
            if (record.tooltipPreview) {
                nameSpan.title = record.tooltipPreview;
            }

            item.appendChild(orderNum);
            item.appendChild(nameSpan);

            item.addEventListener('click', () => {
                this.selectRecord(record, true);
            });

            this.listboxEl.appendChild(item);
        });
    }

    selectRecord(record, notify = true) {
        this.selectedRecord = record;

        // Update selected class on items
        const allItems = this.listboxEl.querySelectorAll('.win-listbox-item');
        allItems.forEach(item => {
            if (parseInt(item.dataset.rowIndex) === record.rowIndex) {
                item.classList.add('selected');
                item.scrollIntoView({ block: 'nearest' });
            } else {
                item.classList.remove('selected');
            }
        });

        if (notify && this.onRecordSelected) {
            this.onRecordSelected(record);
        }
    }

    navigateRelative(direction) {
        if (this.filteredRecords.length === 0) return;
        const currentIdx = this.filteredRecords.findIndex(r => this.selectedRecord && r.rowIndex === this.selectedRecord.rowIndex);
        let targetIdx = currentIdx + direction;
        if (targetIdx < 0) targetIdx = 0;
        if (targetIdx >= this.filteredRecords.length) targetIdx = this.filteredRecords.length - 1;

        if (targetIdx !== currentIdx) {
            this.selectRecord(this.filteredRecords[targetIdx], true);
        }
    }

    updateRecordTitle(rowIndex, newName) {
        const r = this.records.find(item => item.rowIndex === rowIndex);
        if (r) {
            r.keyDisplay = newName;
        }
        const itemEl = this.listboxEl?.querySelector(`[data-row-index="${rowIndex}"] .item-name-text`);
        if (itemEl) {
            itemEl.textContent = newName || `Dòng #${rowIndex}`;
        }
    }
}
