// Searchable and Dependent Dropdown Component
import { api } from '../api.js';

export class DropdownHandler {
    constructor() {
        this.cache = new Map();
        this.activePopover = null;

        document.addEventListener('click', (e) => {
            if (this.activePopover && !this.activePopover.contains(e.target) && !e.target.closest('.dropdown-trigger')) {
                this.closeActivePopover();
            }
        });
    }

    closeActivePopover() {
        if (this.activePopover) {
            this.activePopover.remove();
            this.activePopover = null;
        }
    }

    async getOptions(columnIndex, parents = null) {
        const cacheKey = parents && Object.keys(parents).length > 0 
            ? `${columnIndex}_${JSON.stringify(parents)}` 
            : `${columnIndex}`;

        if (this.cache.has(cacheKey)) {
            return this.cache.get(cacheKey);
        }

        try {
            let items;
            if (parents && Object.keys(parents).length > 0) {
                items = await api.getDependentDropdown(columnIndex, parents);
            } else {
                items = await api.getDropdown(columnIndex);
            }
            items = items || [];
            this.cache.set(cacheKey, items);
            return items;
        } catch (e) {
            console.error(`Error loading dropdown for col ${columnIndex}:`, e);
            return [];
        }
    }

    clearCache(columnIndex = null) {
        if (columnIndex !== null) {
            for (const key of this.cache.keys()) {
                if (key.startsWith(`${columnIndex}_`) || key === `${columnIndex}`) {
                    this.cache.delete(key);
                }
            }
        } else {
            this.cache.clear();
        }
    }

    // Creates interactive searchable dropdown element
    createSearchableDropdown(field, currentValue, onSelect) {
        const wrapper = document.createElement('div');
        wrapper.className = 'searchable-dropdown-wrapper';

        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-input dropdown-trigger';
        input.value = currentValue || '';
        input.placeholder = field.hasDropdown ? 'Chọn hoặc nhập giá trị...' : '';
        input.autocomplete = 'off';

        const triggerBtn = document.createElement('button');
        triggerBtn.type = 'button';
        triggerBtn.className = 'dropdown-btn-arrow';
        triggerBtn.innerHTML = `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path fill-rule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clip-rule="evenodd" /></svg>`;

        wrapper.appendChild(input);
        wrapper.appendChild(triggerBtn);

        const openDropdown = async () => {
            this.closeActivePopover();

            const popover = document.createElement('div');
            popover.className = 'dropdown-popover glassmorphism-card';

            const searchInput = document.createElement('input');
            searchInput.type = 'text';
            searchInput.className = 'dropdown-search-input';
            searchInput.placeholder = 'Tìm kiếm nhanh...';
            popover.appendChild(searchInput);

            const list = document.createElement('div');
            list.className = 'dropdown-options-list';
            list.innerHTML = '<div class="dropdown-loading">Đang tải danh sách...</div>';
            popover.appendChild(list);

            // Position popover
            document.body.appendChild(popover);
            const rect = wrapper.getBoundingClientRect();
            popover.style.top = `${rect.bottom + window.scrollY + 4}px`;
            popover.style.left = `${rect.left + window.scrollX}px`;
            popover.style.width = `${Math.max(rect.width, 240)}px`;

            this.activePopover = popover;

            const allItems = await this.getOptions(field.columnIndex);
            const renderItems = (filtered) => {
                list.innerHTML = '';
                if (filtered.length === 0) {
                    list.innerHTML = '<div class="dropdown-empty">Không tìm thấy kết quả</div>';
                    return;
                }
                filtered.forEach(item => {
                    const itemEl = document.createElement('div');
                    itemEl.className = 'dropdown-item' + (item === input.value ? ' selected' : '');
                    itemEl.textContent = item;
                    itemEl.addEventListener('click', () => {
                        input.value = item;
                        this.closeActivePopover();
                        onSelect(item);
                    });
                    list.appendChild(itemEl);
                });
            };

            renderItems(allItems);
            searchInput.focus();

            searchInput.addEventListener('input', () => {
                const q = searchInput.value.trim().toLowerCase();
                if (!q) {
                    renderItems(allItems);
                } else {
                    renderItems(allItems.filter(i => i.toLowerCase().includes(q)));
                }
            });

            searchInput.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    this.closeActivePopover();
                    input.focus();
                } else if (e.key === 'Enter') {
                    const first = list.querySelector('.dropdown-item');
                    if (first) {
                        first.click();
                    } else if (searchInput.value.trim()) {
                        input.value = searchInput.value.trim();
                        this.closeActivePopover();
                        onSelect(input.value);
                    }
                }
            });
        };

        triggerBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            if (this.activePopover) {
                this.closeActivePopover();
            } else {
                openDropdown();
            }
        });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown' || (e.altKey && e.key === 'ArrowDown')) {
                e.preventDefault();
                openDropdown();
            }
        });

        input.addEventListener('change', () => {
            onSelect(input.value);
        });

        return wrapper;
    }
}

export const dropdownManager = new DropdownHandler();
