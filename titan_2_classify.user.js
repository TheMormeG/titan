// ==UserScript==
// @name         退换货功能2：分类
// @namespace    http://tampermonkey.net/
// @version      v.2026-04-28
// @description  备注显示；筛选；问题组关联沟通人（支持多选，自动学习）；沟通状态四档（—/未沟通/沟通中/已沟通）；支持图片附件缓存（Blob存储，点击放大）；沟通人智能补全；问题分类智能补全；【优化】日期备注独立一行显示（mm.dd），居中，颜色与沟通状态同步
// @author       TheMorme
// @match        http://203.166.165.221:7000/titan/return-exchange-center!returnExchangeNewList.action
// @icon         https://ts2.tc.mm.bing.net/th/id/ODF.eE5djTSqY3rSuOvRaRkm4g?w=32&h=32&qlt=91&pcl=fffffa&o=6&pid=1.2
// @require      https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // ========== 用户配置区 ==========
    const WAREHOUSE_CATEGORIES = ['空瓶', '过期', '配错', '漏配'];
    const DELIVERY_CATEGORIES = ['破损/碎', '泄露'];
    const FOCUS_CATEGORIES = WAREHOUSE_CATEGORIES;
    const UNCOMMUNICATED_COLOR = '#ff0000';
    const IN_PROGRESS_COLOR = '#fab001';
    const COMMUNICATED_COLOR = '#2f6f2f';
    const DEFAULT_COLOR = '#333333';

    const CONTACT_SEPARATOR = '、';
    const MULTI_CONTACT_DISPLAY_LIMIT = 2;

    const CATEGORY_GROUP_MAP = {
        '仓储': FOCUS_CATEGORIES,
        '配送': DELIVERY_CATEGORIES
    };
    // ===============================

    const STORAGE_PREFIX = 'refund_category_';
    const IMAGE_DB_NAME = 'RefundImagesDB';
    const IMAGE_STORE_NAME = 'images';
    const DB_VERSION = 1;
    const ANCHOR_COLUMN_TEXT = '问题描述';
    const DEFAULT_CATEGORY_WIDTH = '100px';

    let columnAdded = false;
    let currentFilterCategory = '全部';
    let currentFilterStatus = '全部';
    let currentFilterContact = '全部';
    let db = null;

    // ---------- IndexedDB 初始化 ----------
    function initDB() {
        return new Promise((resolve, reject) => {
            if (db) return resolve(db);
            const request = indexedDB.open(IMAGE_DB_NAME, DB_VERSION);
            request.onerror = () => reject(request.error);
            request.onsuccess = () => {
                db = request.result;
                resolve(db);
            };
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(IMAGE_STORE_NAME)) {
                    db.createObjectStore(IMAGE_STORE_NAME, { keyPath: 'id' });
                }
            };
        });
    }

    // 存储图片（Blob 数组）
    async function saveImagesToDB(refundNo, blobs) {
        await initDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IMAGE_STORE_NAME, 'readwrite');
            const store = tx.objectStore(IMAGE_STORE_NAME);
            const data = { id: refundNo, images: blobs };
            const request = store.put(data);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    }

    // 读取图片（返回 Blob 数组）
    async function loadImagesFromDB(refundNo) {
        await initDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IMAGE_STORE_NAME, 'readonly');
            const store = tx.objectStore(IMAGE_STORE_NAME);
            const request = store.get(refundNo);
            request.onsuccess = () => resolve(request.result?.images || []);
            request.onerror = () => reject(request.error);
        });
    }

    // 删除图片记录
    async function deleteImagesFromDB(refundNo) {
        await initDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(IMAGE_STORE_NAME, 'readwrite');
            const store = tx.objectStore(IMAGE_STORE_NAME);
            const request = store.delete(refundNo);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    }

    // ---------- 辅助函数 ----------
    function getGroupForCategory(category) {
        if (!category) return null;
        for (const [group, categories] of Object.entries(CATEGORY_GROUP_MAP)) {
            if (categories.includes(category)) {
                return group;
            }
        }
        return null;
    }

    function getContactsForGroup(groupName) {
        const targetCategories = CATEGORY_GROUP_MAP[groupName];
        if (!targetCategories) return new Set();

        const contacts = new Set();
        for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (!key || !key.startsWith(STORAGE_PREFIX)) continue;
            try {
                const raw = localStorage.getItem(key);
                const parsed = JSON.parse(raw);
                const category = parsed?.category?.trim();
                let contactData = parsed?.contactPerson;

                if (category && targetCategories.includes(category)) {
                    let firstContact = null;
                    if (typeof contactData === 'string') {
                        if (contactData.startsWith('[') && contactData.endsWith(']')) {
                            try {
                                const arr = JSON.parse(contactData);
                                firstContact = Array.isArray(arr) && arr.length > 0 ? arr[0] : contactData;
                            } catch(e) {
                                firstContact = contactData;
                            }
                        } else {
                            firstContact = contactData;
                        }
                    } else if (Array.isArray(contactData) && contactData.length > 0) {
                        firstContact = contactData[0];
                    }
                    if (firstContact && firstContact.trim()) {
                        contacts.add(firstContact.trim());
                    }
                }
            } catch (e) {}
        }
        contacts.delete("[]");
        return contacts;
    }

    function getCurrentCategoryColumnWidth() {
        const headerTable = document.querySelector('.header_list table');
        if (!headerTable) return DEFAULT_CATEGORY_WIDTH;
        const headerRow = headerTable.querySelector('tr.tab_bt') || headerTable.querySelector('tr');
        if (!headerRow) return DEFAULT_CATEGORY_WIDTH;
        for (let i = 0; i < headerRow.cells.length; i++) {
            if (headerRow.cells[i].textContent.trim() === '备注') {
                const width = headerRow.cells[i].offsetWidth;
                if (width > 0) return width + 'px';
                break;
            }
        }
        return DEFAULT_CATEGORY_WIDTH;
    }

    function syncTableColumnWidths() {
        const headerTable = document.querySelector('.header_list table');
        const dataTable = document.querySelector('.inner_list table');
        if (!headerTable || !dataTable) return;

        headerTable.style.tableLayout = 'fixed';
        dataTable.style.tableLayout = 'fixed';

        const headerRow = headerTable.querySelector('tr.tab_bt') || headerTable.querySelector('tr');
        if (!headerRow) return;

        const headerCells = Array.from(headerRow.cells);
        const dataRows = dataTable.querySelectorAll('tr');

        headerCells.forEach((cell, idx) => {
            const width = cell.offsetWidth;
            if (width > 0) {
                cell.style.width = width + 'px';
                dataRows.forEach(row => {
                    if (row.cells[idx]) {
                        row.cells[idx].style.width = width + 'px';
                    }
                });
            }
        });
    }

    function getDefaultData() {
        return {
            isCommunicated: '—',
            category: '',
            contactPerson: [],
            feedback: '',
            dateRemark: ''
        };
    }

    function loadCategory(refundNo) {
        const raw = localStorage.getItem(STORAGE_PREFIX + refundNo);
        if (!raw) return getDefaultData();
        try {
            const parsed = JSON.parse(raw);
            let contact = parsed.contactPerson || '';
            let contactArray = [];
            if (typeof contact === 'string') {
                if (contact.startsWith('[') && contact.endsWith(']')) {
                    try { contactArray = JSON.parse(contact); } catch(e) { contactArray = contact ? [contact] : []; }
                } else {
                    contactArray = contact ? [contact] : [];
                }
            } else if (Array.isArray(contact)) {
                contactArray = contact;
            }
            let isCommunicated = parsed.isCommunicated || '—';

            return {
                isCommunicated: isCommunicated,
                category: parsed.category || '',
                contactPerson: contactArray.filter(c => c && c.trim() !== ''),
                feedback: parsed.feedback || '',
                dateRemark: parsed.dateRemark || ''
            };
        } catch (e) {
            const oldText = raw.trim();
            const newData = getDefaultData();
            if (oldText) newData.feedback = oldText;
            saveCategory(refundNo, newData);
            return newData;
        }
    }

    function saveCategory(refundNo, data) {
        if (!refundNo) return;
        const contactArray = data.contactPerson || [];
        const isEmpty = (!data.category || data.category.trim() === '') &&
                        contactArray.length === 0 &&
                        (!data.feedback || data.feedback.trim() === '') &&
                        (data.isCommunicated === '—'|| !data.isCommunicated) &&
                        (!data.dateRemark || data.dateRemark.trim() === '');
        if (isEmpty) {
            localStorage.removeItem(STORAGE_PREFIX + refundNo);
        } else {
            const toStore = {
                isCommunicated: data.isCommunicated,
                category: data.category || '',
                contactPerson: JSON.stringify(contactArray),
                feedback: data.feedback || '',
                dateRemark: data.dateRemark || ''
            };
            localStorage.setItem(STORAGE_PREFIX + refundNo, JSON.stringify(toStore));
        }
    }

    function getContactDisplay(contactArray) {
        if (!contactArray || contactArray.length === 0) return '—';
        if (contactArray.length === 1) return contactArray[0];
        if (contactArray.length <= MULTI_CONTACT_DISPLAY_LIMIT) return contactArray.join(CONTACT_SEPARATOR);
        return contactArray.slice(0, MULTI_CONTACT_DISPLAY_LIMIT).join(CONTACT_SEPARATOR) + '等';
    }

    // 根据沟通状态获取前缀符号
    function getStatusSymbol(status) {
        if (status === '已沟通') return '✔';
        if (status === '未沟通') return '×';
        if (status === '沟通中') return '*';
        return '';
    }

    // 将 YYYY-MM-DD 转换为 mm.dd
    function formatDateToMMDD(dateStr) {
        if (!dateStr) return '';
        const parts = dateStr.split('-');
        if (parts.length !== 3) return dateStr;
        return `${parts[1]}.${parts[2]}`;
    }

    function getTitleText(data) {
        const contactStr = data.contactPerson.length > 0 ? data.contactPerson.join('、') : '—';
        const dateStr = data.dateRemark ? formatDateToMMDD(data.dateRemark) : '—';
        return `沟通状态: ${data.isCommunicated}\n问题分类: ${data.category || '—'}\n沟通人: ${contactStr}\n沟通反馈: ${data.feedback || '—'}\n备注日期: ${dateStr}`;
    }

    // 颜色逻辑：根据沟通状态决定，无状态时使用默认色
    function getColorByData(data) {
        const status = data.isCommunicated;
        if (status === '已沟通') return COMMUNICATED_COLOR;
        if (status === '未沟通') return UNCOMMUNICATED_COLOR;
        if (status === '沟通中') return IN_PROGRESS_COLOR;
        return DEFAULT_COLOR;
    }

    function getRefundNoFromRow(row) {
        const firstTd = row.querySelector('td:first-child');
        if (!firstTd) return '';
        const span = firstTd.querySelector('span');
        if (span) {
            const title = span.getAttribute('title');
            if (title) return title;
            return span.textContent.trim();
        }
        return firstTd.textContent.trim();
    }

    function buildCategoryCell(refundNo, columnWidth = null) {
        const data = loadCategory(refundNo);
        const titleText = getTitleText(data);
        const color = getColorByData(data);
        const width = columnWidth || getCurrentCategoryColumnWidth();

        const td = document.createElement('td');
        td.style.width = width;
        td.style.minWidth = width;
        td.style.maxWidth = width;
        td.style.padding = '4px 2px';
        td.style.verticalAlign = 'middle';

        const container = document.createElement('div');
        container.style.display = 'flex';
        container.style.alignItems = 'center';
        container.style.gap = '4px';

        // 左侧内容：两行居中显示
        const leftDiv = document.createElement('div');
        leftDiv.style.flex = '1';
        leftDiv.style.display = 'flex';
        leftDiv.style.flexDirection = 'column';
        leftDiv.style.alignItems = 'center';
        leftDiv.style.justifyContent = 'center';
        leftDiv.style.textAlign = 'center';
        leftDiv.style.color = color;
        leftDiv.style.fontWeight = '500';
        leftDiv.title = titleText;

        // 第一行：状态符号 + 问题分类
        const firstLine = document.createElement('div');
        const symbol = getStatusSymbol(data.isCommunicated);
        const cat = data.category && data.category.trim() !== '' ? data.category.trim() : '—';
        firstLine.textContent = symbol ? `${symbol} ${cat}` : cat;
        leftDiv.appendChild(firstLine);

        // 第二行：日期 (mm.dd)
        if (data.dateRemark && data.dateRemark.trim() !== '') {
            const secondLine = document.createElement('div');
            secondLine.textContent = formatDateToMMDD(data.dateRemark);
            secondLine.style.fontSize = '11px';
            secondLine.style.marginTop = '1';
            secondLine.style.lineHeight = '1.2';  // 可选，进一步压缩
            leftDiv.appendChild(secondLine);
        }

        // 编辑按钮
        const btn = document.createElement('button');
        btn.textContent = '✎';
        btn.style.cursor = 'pointer';
        btn.style.border = '1px solid #aaa';
        btn.style.background = '#fff';
        btn.style.borderRadius = '3px';
        btn.style.padding = '2px 5px';
        btn.style.fontSize = '11px';
        btn.className = 'edit-category-btn';
        btn.type = 'button';
        btn.setAttribute('data-refundno', refundNo);

        container.appendChild(leftDiv);
        container.appendChild(btn);
        td.appendChild(container);
        return td;
    }

    function refreshRowCategoryCell(refundNo) {
        if (!refundNo) return;
        const dataTable = document.querySelector('.inner_list table');
        if (!dataTable) return;
        const rows = dataTable.querySelectorAll('tbody tr');
        let targetRow = null;
        for (let row of rows) {
            if (getRefundNoFromRow(row) === refundNo) {
                targetRow = row;
                break;
            }
        }
        if (!targetRow) return;

        const headerTable = document.querySelector('.header_list table');
        if (!headerTable) return;
        const headerRow = headerTable.querySelector('tr.tab_bt') || headerTable.querySelector('tr');
        let categoryColIndex = -1;
        for (let i = 0; i < headerRow.cells.length; i++) {
            if (headerRow.cells[i].textContent.trim() === '备注') {
                categoryColIndex = i;
                break;
            }
        }
        if (categoryColIndex === -1) return;

        const currentWidth = getCurrentCategoryColumnWidth();
        const newTd = buildCategoryCell(refundNo, currentWidth);
        const oldTd = targetRow.cells[categoryColIndex];
        if (oldTd) oldTd.replaceWith(newTd);
        else {
            if (categoryColIndex >= targetRow.cells.length) targetRow.appendChild(newTd);
            else targetRow.insertBefore(newTd, targetRow.cells[categoryColIndex]);
        }
    }

    function refreshAllRows() {
        const dataTable = document.querySelector('.inner_list table');
        if (!dataTable) return;
        const rows = dataTable.querySelectorAll('tbody tr');
        if (rows.length === 0) return;
        const headerTable = document.querySelector('.header_list table');
        if (!headerTable) return;
        const headerRow = headerTable.querySelector('tr.tab_bt') || headerTable.querySelector('tr');
        let categoryColIndex = -1;
        for (let i = 0; i < headerRow.cells.length; i++) {
            if (headerRow.cells[i].textContent.trim() === '备注') {
                categoryColIndex = i;
                break;
            }
        }
        if (categoryColIndex === -1) return;

        const currentWidth = getCurrentCategoryColumnWidth();
        rows.forEach(row => {
            const refundNo = getRefundNoFromRow(row);
            if (!refundNo) return;
            const newTd = buildCategoryCell(refundNo, currentWidth);
            const oldTd = row.cells[categoryColIndex];
            if (oldTd) oldTd.replaceWith(newTd);
            else {
                if (categoryColIndex >= row.cells.length) row.appendChild(newTd);
                else row.insertBefore(newTd, row.cells[categoryColIndex]);
            }
        });
    }

    // ---------- 筛选功能 ----------
    function collectCategoryOptions() {
        const dataTable = document.querySelector('.inner_list table');
        if (!dataTable) return ['全部'];
        const rows = dataTable.querySelectorAll('tbody tr');
        const categories = new Set();
        rows.forEach(row => {
            const refundNo = getRefundNoFromRow(row);
            if (refundNo) {
                const data = loadCategory(refundNo);
                if (data.category && data.category.trim() !== '') {
                    categories.add(data.category.trim());
                }
            }
        });
        const sorted = Array.from(categories).sort();
        return ['全部', ...sorted];
    }

    function collectContactOptions() {
        const dataTable = document.querySelector('.inner_list table');
        if (!dataTable) return ['全部'];
        const rows = dataTable.querySelectorAll('tbody tr');
        const contacts = new Set();
        rows.forEach(row => {
            const refundNo = getRefundNoFromRow(row);
            if (refundNo) {
                const data = loadCategory(refundNo);
                data.contactPerson.forEach(c => {
                    if (c && c.trim() !== '') contacts.add(c.trim());
                });
            }
        });
        const sorted = Array.from(contacts).sort((a, b) => a.localeCompare(b));
        return ['全部', ...sorted];
    }

    function refreshFilterOptions() {
        const categorySelect = document.getElementById('categoryFilterSelect');
        const statusSelect = document.getElementById('statusFilterSelect');
        const contactSelect = document.getElementById('contactFilterSelect');
        if (!categorySelect || !statusSelect || !contactSelect) return;

        const categoryOptions = collectCategoryOptions();
        const currentCategory = categorySelect.value;
        categorySelect.innerHTML = '';
        categoryOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            categorySelect.appendChild(option);
        });
        if (categoryOptions.includes(currentCategory)) {
            categorySelect.value = currentCategory;
        } else {
            categorySelect.value = '全部';
        }

        const contactOptions = collectContactOptions();
        const currentContact = contactSelect.value;
        contactSelect.innerHTML = '';
        contactOptions.forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            contactSelect.appendChild(option);
        });
        if (contactOptions.includes(currentContact)) {
            contactSelect.value = currentContact;
        } else {
            contactSelect.value = '全部';
        }

        currentFilterCategory = categorySelect.value;
        currentFilterStatus = statusSelect.value;
        currentFilterContact = contactSelect.value;
    }

    function applyFilter() {
        const categorySelect = document.getElementById('categoryFilterSelect');
        const statusSelect = document.getElementById('statusFilterSelect');
        const contactSelect = document.getElementById('contactFilterSelect');
        const dateFilterValue = document.getElementById('dateRemarkFilter')?.value || '';
        if (!categorySelect || !statusSelect || !contactSelect) return;

        const selectedCategory = categorySelect.value;
        const selectedStatus = statusSelect.value;
        const selectedContact = contactSelect.value;
        currentFilterCategory = selectedCategory;
        currentFilterStatus = selectedStatus;
        currentFilterContact = selectedContact;

        const dataTable = document.querySelector('.inner_list table');
        if (!dataTable) return;
        const rows = dataTable.querySelectorAll('tbody tr');
        if (rows.length === 0) return;

        rows.forEach(row => {
            const refundNo = getRefundNoFromRow(row);
            if (!refundNo) {
                row.style.display = '';
                return;
            }
            const data = loadCategory(refundNo);
            const category = data.category ? data.category.trim() : '';
            const status = data.isCommunicated;
            const contacts = data.contactPerson;

            let categoryMatch = (selectedCategory === '全部' || category === selectedCategory);
            let statusMatch = false;
            if (selectedStatus === '全部') {
                statusMatch = true;
            } else {
                if (category === '') {
                    statusMatch = false;
                } else {
                    statusMatch = (status === selectedStatus);
                }
            }
            let contactMatch = (selectedContact === '全部' || contacts.includes(selectedContact));

            let dateMatch = true;
            if (dateFilterValue) {
                const rowDate = data.dateRemark || '';
                dateMatch = (rowDate === dateFilterValue);
            }

            row.style.display = (categoryMatch && statusMatch && contactMatch && dateMatch) ? '' : 'none';
        });
    }

    function addFilterControl() {
        if (document.getElementById('categoryFilterContainer')) return;

        const exportBtn = document.getElementById('custom_export_all_btn');
        let parentContainer = null;
        if (exportBtn && exportBtn.parentElement) {
            parentContainer = exportBtn.parentElement;
        } else {
            const tabBar = document.querySelector('.exchange_tab .tab_bar');
            if (tabBar) parentContainer = tabBar;
        }
        if (!parentContainer) return;

        const filterContainer = document.createElement('div');
        filterContainer.id = 'categoryFilterContainer';
        filterContainer.style.display = 'inline-block';
        filterContainer.style.marginLeft = '15px';
        filterContainer.style.verticalAlign = 'middle';

        const statusLabel = document.createElement('span');
        statusLabel.textContent = '沟通状态：';
        statusLabel.style.marginRight = '5px';
        statusLabel.style.fontSize = '12px';

        const statusSelect = document.createElement('select');
        statusSelect.id = 'statusFilterSelect';
        statusSelect.style.padding = '4px 8px';
        statusSelect.style.borderRadius = '4px';
        statusSelect.style.border = '1px solid #ccc';
        statusSelect.style.fontSize = '12px';
        statusSelect.style.marginRight = '10px';
        ['全部', '—', '未沟通', '沟通中', '已沟通'].forEach(opt => {
            const option = document.createElement('option');
            option.value = opt;
            option.textContent = opt;
            statusSelect.appendChild(option);
        });

        const categoryLabel = document.createElement('span');
        categoryLabel.textContent = '问题分类：';
        categoryLabel.style.marginRight = '5px';
        categoryLabel.style.fontSize = '12px';

        const categorySelect = document.createElement('select');
        categorySelect.id = 'categoryFilterSelect';
        categorySelect.style.padding = '4px 8px';
        categorySelect.style.borderRadius = '4px';
        categorySelect.style.border = '1px solid #ccc';
        categorySelect.style.fontSize = '12px';
        categorySelect.style.marginRight = '10px';

        const contactLabel = document.createElement('span');
        contactLabel.textContent = '沟通人：';
        contactLabel.style.marginRight = '5px';
        contactLabel.style.fontSize = '12px';

        const contactSelect = document.createElement('select');
        contactSelect.id = 'contactFilterSelect';
        contactSelect.style.padding = '4px 8px';
        contactSelect.style.borderRadius = '4px';
        contactSelect.style.border = '1px solid #ccc';
        contactSelect.style.fontSize = '12px';
        contactSelect.style.marginRight = '10px';

        const dateLabel = document.createElement('span');
        dateLabel.textContent = '备注日期：';
        dateLabel.style.marginRight = '5px';
        dateLabel.style.fontSize = '12px';
        
        const dateInput = document.createElement('input');
        dateInput.type = 'date';
        dateInput.id = 'dateRemarkFilter';
        dateInput.style.padding = '4px 8px';
        dateInput.style.borderRadius = '4px';
        dateInput.style.border = '1px solid #ccc';
        dateInput.style.fontSize = '12px';
        dateInput.style.marginRight = '5px';
        
        const clearDateBtn = document.createElement('button');
        clearDateBtn.textContent = '清除';
        clearDateBtn.style.padding = '4px 8px';
        clearDateBtn.style.cursor = 'pointer';
        clearDateBtn.style.border = '1px solid #dcdfe6';
        clearDateBtn.style.background = '#fff';
        clearDateBtn.style.borderRadius = '4px';
        clearDateBtn.style.fontSize = '12px';
        
        filterContainer.appendChild(dateLabel);
        filterContainer.appendChild(dateInput);
        filterContainer.appendChild(clearDateBtn);
                
        const refreshBtn = document.createElement('button');
        refreshBtn.textContent = '刷新列表';
        refreshBtn.style.padding = '4px 8px';
        refreshBtn.style.cursor = 'pointer';
        refreshBtn.style.border = '1px solid #409eff';
        refreshBtn.style.background = '#fff';
        refreshBtn.style.color = '#409eff';
        refreshBtn.style.borderRadius = '4px';
        refreshBtn.style.fontSize = '12px';

        filterContainer.appendChild(statusLabel);
        filterContainer.appendChild(statusSelect);
        filterContainer.appendChild(categoryLabel);
        filterContainer.appendChild(categorySelect);
        filterContainer.appendChild(contactLabel);
        filterContainer.appendChild(contactSelect);
        filterContainer.appendChild(refreshBtn);
        parentContainer.appendChild(filterContainer);

        statusSelect.addEventListener('change', applyFilter);
        categorySelect.addEventListener('change', applyFilter);
        contactSelect.addEventListener('change', applyFilter);
        dateInput.addEventListener('change', applyFilter);
        refreshBtn.addEventListener('click', () => {
            refreshFilterOptions();
            applyFilter();
        });

        refreshFilterOptions();
    }

    // ---------- 历史记录收集（用于自动补全）----------
    function collectHistorySuggestions() {
        const categories = new Set();
        const contactPersons = new Set();
        for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key && key.startsWith(STORAGE_PREFIX)) {
                const data = loadCategory(key.slice(STORAGE_PREFIX.length));
                if (data.category && data.category.trim() !== '') {
                    categories.add(data.category.trim());
                }
                data.contactPerson.forEach(c => {
                    const trimmed = c?.trim?.() || '';
                    if (trimmed && trimmed !== '[]' && !trimmed.startsWith('[')) {
                        contactPersons.add(trimmed);
                    }
                });
            }
        }
        // 预设分类：合并仓库和配送分类，并加上“有效期短”
        const preset = [...new Set([...WAREHOUSE_CATEGORIES, ...DELIVERY_CATEGORIES, '有效期短'])];
        preset.forEach(c => categories.add(c));
        return {
            categories: Array.from(categories).sort((a,b)=>a.localeCompare(b)),
            contactPersons: Array.from(contactPersons).sort((a, b) => a.localeCompare(b))
        };
    }

    // 缓存全量联系人列表（用于有输入时的匹配）
    let fullContactListCache = null;
    function getFullContactList() {
        if (!fullContactListCache) {
            fullContactListCache = collectHistorySuggestions().contactPersons;
        }
        return fullContactListCache;
    }

    // ---------- 自定义下拉组件（支持单选/多选模式）----------
    class CustomDropdown {
        constructor(inputElement, suggestionGetter, options = { multiSelect: true }) {
            this.input = inputElement;
            this.getSuggestions = suggestionGetter;
            this.multiSelect = options.multiSelect;
            this.panel = null;
            this.items = [];
            this.selectedIndex = -1;
            this.isOpen = false;
            this.separator = CONTACT_SEPARATOR;
            this.initEventListeners();
        }

        getLastSegment(fullValue) {
            if (!this.multiSelect) {
                return { prefix: '', last: fullValue || '' };
            }
            if (!fullValue) return { prefix: '', last: '' };
            const parts = fullValue.split(this.separator);
            const lastPart = parts[parts.length - 1].trim();
            if (parts.length === 1) {
                return { prefix: '', last: lastPart };
            } else {
                const prefix = parts.slice(0, -1).join(this.separator) + this.separator;
                return { prefix, last: lastPart };
            }
        }

        isMultiSelectMode(fullValue) {
            if (!this.multiSelect) return false;
            if (!fullValue) return false;
            const parts = fullValue.split(this.separator);
            if (parts.length < 2) return false;
            const lastPart = parts[parts.length - 1].trim();
            return lastPart.length > 0;
        }

        filterSuggestionsByLastSegment(suggestions, lastSegment) {
            if (!lastSegment) return suggestions;
            const lowerLast = lastSegment.toLowerCase();
            const filtered = suggestions.filter(item => item.toLowerCase().includes(lowerLast));
            filtered.sort((a, b) => {
                const aLower = a.toLowerCase();
                const bLower = b.toLowerCase();
                if (aLower === lowerLast && bLower !== lowerLast) return -1;
                if (bLower === lowerLast && aLower !== lowerLast) return 1;
                if (aLower.startsWith(lowerLast) && !bLower.startsWith(lowerLast)) return -1;
                if (bLower.startsWith(lowerLast) && !aLower.startsWith(lowerLast)) return 1;
                return a.localeCompare(b);
            });
            return filtered;
        }

        computeNewValue(oldValue, selectedItem) {
            const trimmedSelected = selectedItem.trim();
            if (!trimmedSelected) return oldValue;

            if (!this.multiSelect) {
                return trimmedSelected;
            }

            const fullValue = oldValue || '';

            // 情况1：已经处于“多段输入”模式（有分隔符且最后一段非空）
            if (this.isMultiSelectMode(fullValue)) {
                const { prefix } = this.getLastSegment(fullValue);
                const newValue = prefix + trimmedSelected;
                const newParts = newValue.split(this.separator).map(s => s.trim()).filter(s => s);
                const uniqueParts = [...new Set(newParts)];
                return uniqueParts.join(this.separator);
            }

            // 情况2：输入为空 → 直接设置
            if (fullValue === '') {
                return trimmedSelected;
            }

            // 情况3：以分隔符结尾（例如“张三、”） → 追加选中项
            if (fullValue.endsWith(this.separator)) {
                return fullValue + trimmedSelected;
            }

            // 情况4：包含分隔符但未以分隔符结尾（理论上不会触发，但兜底）
            if (fullValue.includes(this.separator)) {
                const lastSepIndex = fullValue.lastIndexOf(this.separator);
                const prefix = fullValue.substring(0, lastSepIndex + 1);
                return prefix + trimmedSelected;
            }

            // 情况5：无分隔符且非空（如“上海”） → 直接替换
            return trimmedSelected;
        }

        updateInputValue(selectedItem) {
            const oldVal = this.input.value;
            const newVal = this.computeNewValue(oldVal, selectedItem);
            this.input.value = newVal;
            this.input.dispatchEvent(new Event('input', { bubbles: true }));
            this.hide();
            this.input.focus();
        }

        getMatchScore(item, inputValue) {
            if (!inputValue) return 0;
            const lowerItem = item.toLowerCase();
            const lowerInput = inputValue.toLowerCase();
            if (lowerItem === lowerInput) return 3;
            if (lowerItem.startsWith(lowerInput)) return 2;
            if (lowerItem.includes(lowerInput)) return 1;
            return 0;
        }

        initEventListeners() {
            this.input.addEventListener('input', () => this.show());
            this.input.addEventListener('focus', () => this.show());
            this.input.addEventListener('keydown', (e) => this.handleKeydown(e));

            document.addEventListener('click', (e) => {
                if (!this.input.contains(e.target) && this.panel && !this.panel.contains(e.target)) {
                    this.hide();
                }
            });

            window.addEventListener('scroll', (e) => {
                if (this.panel && this.panel.contains(e.target)) return;
                this.hide();
            }, true);
        }

        createPanel() {
            const panel = document.createElement('div');
            panel.className = 'custom-dropdown-panel';
            panel.style.cssText = `
                position: absolute;
                background: white;
                border: 1px solid #ccc;
                border-radius: 6px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.15);
                max-height: 280px;
                overflow-y: auto;
                z-index: 10001;
                display: none;
                min-width: 200px;
                font-size: 13px;
            `;
            document.body.appendChild(panel);
            return panel;
        }

        destroyPanel() {
            if (this.panel && this.panel.parentNode) {
                this.panel.parentNode.removeChild(this.panel);
            }
            this.panel = null;
        }

        show() {
            this.destroyPanel();
            this.panel = this.createPanel();

            let rawSuggestions = this.getSuggestions();
            if (!Array.isArray(rawSuggestions)) rawSuggestions = [];
            const fullValue = this.input.value;
            const isMulti = this.multiSelect && this.isMultiSelectMode(fullValue);
            let displayItems = [...rawSuggestions];

            if (isMulti) {
                const { last } = this.getLastSegment(fullValue);
                if (last) {
                    displayItems = this.filterSuggestionsByLastSegment(rawSuggestions, last);
                    if (displayItems.length === 1 && displayItems[0] === last) {
                        this.hide();
                        return;
                    }
                }
                if (displayItems.length === 0) {
                    this.hide();
                    return;
                }
            } else {
                const inputVal = fullValue.trim();
                if (inputVal) {
                    displayItems.sort((a, b) => {
                        const scoreA = this.getMatchScore(a, inputVal);
                        const scoreB = this.getMatchScore(b, inputVal);
                        if (scoreA !== scoreB) return scoreB - scoreA;
                        return a.localeCompare(b);
                    });
                    if (displayItems.length === 1 && displayItems[0] === inputVal) {
                        this.hide();
                        return;
                    }
                } else {
                    displayItems.sort((a, b) => a.localeCompare(b));
                }
            }

            if (displayItems.length === 0) {
                this.hide();
                return;
            }

            this.items = displayItems;
            this.renderPanel(displayItems);
            this.positionPanel();
            this.panel.style.display = 'block';
            this.isOpen = true;
            this.selectedIndex = -1;
        }

        renderPanel(items) {
            if (!this.panel) return;
            this.panel.innerHTML = '';
            items.forEach((item, idx) => {
                const div = document.createElement('div');
                div.textContent = item;
                div.style.padding = '8px 12px';
                div.style.cursor = 'pointer';
                div.style.borderBottom = '1px solid #eee';
                div.style.fontSize = '13px';
                div.addEventListener('mouseenter', () => {
                    this.clearSelectedHighlight();
                    this.selectedIndex = idx;
                    div.style.backgroundColor = '#f0f7ff';
                });
                div.addEventListener('mouseleave', () => {
                    if (this.selectedIndex === idx) {
                        div.style.backgroundColor = '';
                    }
                });
                div.addEventListener('click', () => {
                    this.updateInputValue(item);
                });
                this.panel.appendChild(div);
            });
        }

        clearSelectedHighlight() {
            if (this.panel && this.selectedIndex >= 0 && this.panel.children[this.selectedIndex]) {
                this.panel.children[this.selectedIndex].style.backgroundColor = '';
            }
        }

        positionPanel() {
            if (!this.panel) return;
            const rect = this.input.getBoundingClientRect();
            this.panel.style.left = rect.left + 'px';
            this.panel.style.top = (rect.bottom + window.scrollY) + 'px';
            this.panel.style.width = rect.width + 'px';
        }

        hide() {
            this.isOpen = false;
            this.selectedIndex = -1;
            this.destroyPanel();
        }

        handleKeydown(e) {
            if (!this.isOpen) {
                if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                    this.show();
                    e.preventDefault();
                }
                return;
            }
            switch (e.key) {
                case 'ArrowDown':
                    e.preventDefault();
                    this.moveSelection(1);
                    break;
                case 'ArrowUp':
                    e.preventDefault();
                    this.moveSelection(-1);
                    break;
                case 'Enter':
                    e.preventDefault();
                    if (this.selectedIndex >= 0 && this.items[this.selectedIndex]) {
                        const selectedItem = this.items[this.selectedIndex];
                        this.updateInputValue(selectedItem);
                    }
                    break;
                case 'Escape':
                    this.hide();
                    break;
                default:
                    break;
            }
        }

        moveSelection(delta) {
            if (!this.panel) return;
            const newIndex = this.selectedIndex + delta;
            if (newIndex >= 0 && newIndex < this.items.length) {
                this.clearSelectedHighlight();
                this.selectedIndex = newIndex;
                const selectedDiv = this.panel.children[this.selectedIndex];
                selectedDiv.style.backgroundColor = '#f0f7ff';
                selectedDiv.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
            }
        }
    }

    // ---------- 模态框（Blob图片附件，包含日期备注）----------
    let modalInstance = null;
    let categoryDropdown = null;
    let contactDropdown = null;
    let currentImageFiles = [];
    let currentObjectURLs = [];

    function revokeAllObjectURLs() {
        currentObjectURLs.forEach(url => URL.revokeObjectURL(url));
        currentObjectURLs = [];
    }

    function createModal() {
        if (modalInstance) return modalInstance;
        const modalDiv = document.createElement('div');
        modalDiv.id = 'customCategoryModal';
        modalDiv.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 10000;
            font-family: system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif;
        `;
        const modalContent = document.createElement('div');
        modalContent.style.cssText = `
            background: #fff;
            border-radius: 12px;
            width: 550px;
            max-width: 95%;
            padding: 20px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.2);
            max-height: 90vh;
            overflow-y: auto;
        `;
        modalContent.innerHTML = `
            <h3 style="margin-top:0; margin-bottom:16px;">编辑备注信息</h3>
            <div style="margin-bottom:12px;">
                <label style="display:block; margin-bottom:4px; font-weight:500;">沟通状态</label>
                <select id="modalIsCommunicated" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:6px;">
                    <option value="—">—</option>
                    <option value="未沟通">未沟通</option>
                    <option value="沟通中">沟通中</option>
                    <option value="已沟通">已沟通</option>
                </select>
            </div>
            <div style="margin-bottom:12px; position:relative;">
                <label style="display:block; margin-bottom:4px; font-weight:500;">问题分类</label>
                <input type="text" id="modalCategory" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:6px; box-sizing:border-box;" placeholder="例如：空瓶、过期、配错...">
            </div>
            <div style="margin-bottom:12px; position:relative;">
                <label style="display:block; margin-bottom:4px; font-weight:500;">沟通人（支持多选，用、分隔）</label>
                <input type="text" id="modalContactPerson" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:6px; box-sizing:border-box;" placeholder="例如：张专员、李主管">
            </div>
            <div style="margin-bottom:12px;">
                <label style="display:block; margin-bottom:4px; font-weight:500;">备注日期</label>
                <input type="date" id="modalDateRemark" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:6px; box-sizing:border-box;">
            </div>
            <div style="margin-bottom:12px;">
                <label style="display:block; margin-bottom:4px; font-weight:500;">沟通反馈</label>
                <textarea id="modalFeedback" rows="3" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:6px; box-sizing:border-box;" placeholder="详细反馈内容..."></textarea>
            </div>
            <div style="margin-bottom:20px;">
                <label style="display:block; margin-bottom:6px; font-weight:500;">图片附件</label>
                <div style="display:flex; gap:8px; align-items:center; margin-bottom:8px;">
                    <button type="button" id="addImageBtn" style="padding:6px 12px; background:#f0f0f0; border:1px solid #ccc; border-radius:6px; cursor:pointer;">添加图片</button>
                    <span style="font-size:12px; color:#666;">支持多选，图片将自动缓存</span>
                </div>
                <input type="file" id="imageFileInput" accept="image/*" multiple style="display:none;">
                <div id="imagePreviewContainer" style="display:flex; flex-wrap:wrap; gap:10px;"></div>
            </div>
            <div style="display:flex; justify-content:flex-end; gap:10px;">
                <button id="modalCancelBtn" style="padding:6px 12px; background:#f0f0f0; border:none; border-radius:6px; cursor:pointer;">取消</button>
                <button id="modalSaveBtn" style="padding:6px 12px; background:#409eff; color:white; border:none; border-radius:6px; cursor:pointer;">保存</button>
            </div>
        `;
        modalDiv.appendChild(modalContent);
        document.body.appendChild(modalDiv);

        let currentRefundNo = null;
        const hideModal = () => {
            modalDiv.style.display = 'none';
            currentRefundNo = null;
            currentImageFiles = [];
            revokeAllObjectURLs();
            document.getElementById('imagePreviewContainer').innerHTML = '';
            if (categoryDropdown) categoryDropdown.hide();
            if (contactDropdown) contactDropdown.hide();
        };

        const showModal = async (refundNo) => {
            currentRefundNo = refundNo;
            const data = loadCategory(refundNo);
            document.getElementById('modalIsCommunicated').value = data.isCommunicated;
            document.getElementById('modalCategory').value = data.category;
            document.getElementById('modalContactPerson').value = data.contactPerson.join(CONTACT_SEPARATOR);
            document.getElementById('modalDateRemark').value = data.dateRemark || '';
            document.getElementById('modalFeedback').value = data.feedback;

            try {
                const savedBlobs = await loadImagesFromDB(refundNo);
                currentImageFiles = savedBlobs;
                renderImagePreviews(currentImageFiles);
            } catch (e) {
                console.error('加载图片失败', e);
                currentImageFiles = [];
            }

            modalDiv.style.display = 'flex';

            const categoryInput = document.getElementById('modalCategory');
            const contactInput = document.getElementById('modalContactPerson');

            const defaultCategories = [...new Set([...WAREHOUSE_CATEGORIES, ...DELIVERY_CATEGORIES, '有效期短'])].sort((a,b)=>a.localeCompare(b));
            const getAllCategories = () => collectHistorySuggestions().categories;
            const categorySuggestionGetter = () => {
                const inputVal = categoryInput.value.trim();
                if (!inputVal) return defaultCategories;
                const allCategories = getAllCategories();
                const lowerInput = inputVal.toLowerCase();
                const filtered = allCategories.filter(cat => cat.toLowerCase().includes(lowerInput));
                filtered.sort((a, b) => {
                    const aLower = a.toLowerCase();
                    const bLower = b.toLowerCase();
                    if (aLower === lowerInput && bLower !== lowerInput) return -1;
                    if (bLower === lowerInput && aLower !== lowerInput) return 1;
                    if (aLower.startsWith(lowerInput) && !bLower.startsWith(lowerInput)) return -1;
                    if (bLower.startsWith(lowerInput) && !aLower.startsWith(lowerInput)) return 1;
                    return a.localeCompare(b);
                });
                return filtered;
            };
            if (!categoryDropdown) {
                categoryDropdown = new CustomDropdown(categoryInput, categorySuggestionGetter, { multiSelect: false });
            } else {
                categoryDropdown.getSuggestions = categorySuggestionGetter;
            }

            const getDynamicContactSuggestions = () => {
                const currentCategory = categoryInput.value.trim();
                const group = getGroupForCategory(currentCategory);
                const contactFullValue = contactInput.value;
                const { last: lastSegment } = (() => {
                    if (!contactFullValue) return { last: '' };
                    const parts = contactFullValue.split(CONTACT_SEPARATOR);
                    const lastPart = parts[parts.length - 1].trim();
                    return { last: lastPart };
                })();

                if (!lastSegment) {
                    if (group) {
                        const contactsSet = getContactsForGroup(group);
                        if (contactsSet.size > 0) {
                            return [...contactsSet].sort((a, b) => a.localeCompare(b));
                        }
                    }
                    return [];
                } else {
                    const fullList = getFullContactList();
                    const filtered = fullList.filter(contact => contact.toLowerCase().includes(lastSegment.toLowerCase()));
                    filtered.sort((a, b) => {
                        const aLower = a.toLowerCase();
                        const bLower = b.toLowerCase();
                        const lastLower = lastSegment.toLowerCase();
                        if (aLower === lastLower && bLower !== lastLower) return -1;
                        if (bLower === lastLower && aLower !== lastLower) return 1;
                        if (aLower.startsWith(lastLower) && !bLower.startsWith(lastLower)) return -1;
                        if (bLower.startsWith(lastLower) && !aLower.startsWith(lastLower)) return 1;
                        return a.localeCompare(b);
                    });
                    return filtered;
                }
            };

            if (contactDropdown) {
                contactDropdown.getSuggestions = getDynamicContactSuggestions;
            } else {
                contactDropdown = new CustomDropdown(contactInput, getDynamicContactSuggestions, { multiSelect: true });
            }

            const onCategoryChange = () => {
                if (contactDropdown) {
                    const contactFullValue = contactInput.value;
                    const { last } = (() => {
                        if (!contactFullValue) return { last: '' };
                        const parts = contactFullValue.split(CONTACT_SEPARATOR);
                        const lastPart = parts[parts.length - 1].trim();
                        return { last: lastPart };
                    })();
                    if (!last && contactDropdown.isOpen) {
                        contactDropdown.hide();
                        setTimeout(() => {
                            contactDropdown.show();
                        }, 50);
                    } else if (!last && document.activeElement === contactInput) {
                        contactDropdown.show();
                    }
                }
            };
            categoryInput.removeEventListener('input', onCategoryChange);
            categoryInput.removeEventListener('change', onCategoryChange);
            categoryInput.addEventListener('input', onCategoryChange);
            categoryInput.addEventListener('change', onCategoryChange);
        };

        function renderImagePreviews(blobs) {
            const container = document.getElementById('imagePreviewContainer');
            revokeAllObjectURLs();
            container.innerHTML = '';

            let viewer = document.getElementById('imageViewerOverlay');
            if (!viewer) {
                viewer = document.createElement('div');
                viewer.id = 'imageViewerOverlay';
                viewer.style.cssText = `
                    position: fixed;
                    top: 0; left: 0; width: 100%; height: 100%;
                    background: rgba(0,0,0,0.85);
                    display: none;
                    justify-content: center;
                    align-items: center;
                    z-index: 100000;
                    cursor: zoom-out;
                `;
                const viewerImg = document.createElement('img');
                viewerImg.id = 'viewerImage';
                viewerImg.style.cssText = `
                    max-width: 90%;
                    max-height: 90%;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.5);
                    border-radius: 8px;
                    cursor: default;
                `;
                viewer.appendChild(viewerImg);
                viewer.addEventListener('click', () => {
                    viewer.style.display = 'none';
                });
                viewerImg.addEventListener('click', (e) => e.stopPropagation());
                document.body.appendChild(viewer);
            }

            const viewerImg = document.getElementById('viewerImage');

            blobs.forEach((blob, index) => {
                const url = URL.createObjectURL(blob);
                currentObjectURLs.push(url);

                const wrapper = document.createElement('div');
                wrapper.style.position = 'relative';
                wrapper.style.width = '80px';
                wrapper.style.height = '80px';
                wrapper.style.border = '1px solid #ddd';
                wrapper.style.borderRadius = '6px';
                wrapper.style.overflow = 'hidden';
                wrapper.style.cursor = 'zoom-in';

                const img = document.createElement('img');
                img.src = url;
                img.style.width = '100%';
                img.style.height = '100%';
                img.style.objectFit = 'cover';
                img.style.pointerEvents = 'none';

                wrapper.addEventListener('click', (e) => {
                    if (e.target.tagName === 'BUTTON') return;
                    viewerImg.src = url;
                    viewer.style.display = 'flex';
                });

                const delBtn = document.createElement('button');
                delBtn.textContent = '×';
                delBtn.style.position = 'absolute';
                delBtn.style.top = '2px';
                delBtn.style.right = '2px';
                delBtn.style.background = 'rgba(0,0,0,0.6)';
                delBtn.style.color = 'white';
                delBtn.style.border = 'none';
                delBtn.style.borderRadius = '12px';
                delBtn.style.width = '20px';
                delBtn.style.height = '20px';
                delBtn.style.cursor = 'pointer';
                delBtn.style.fontSize = '14px';
                delBtn.style.lineHeight = '18px';
                delBtn.style.padding = '0';
                delBtn.style.zIndex = '2';
                delBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    currentImageFiles.splice(index, 1);
                    renderImagePreviews(currentImageFiles);
                });

                wrapper.appendChild(img);
                wrapper.appendChild(delBtn);
                container.appendChild(wrapper);
            });
        }

        document.getElementById('addImageBtn').addEventListener('click', () => {
            document.getElementById('imageFileInput').click();
        });

        document.getElementById('imageFileInput').addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length === 0) return;
            currentImageFiles.push(...files);
            renderImagePreviews(currentImageFiles);
            e.target.value = '';
        });

        const saveHandler = async () => {
            if (!currentRefundNo) { hideModal(); return; }
            const contactRaw = document.getElementById('modalContactPerson').value.trim();
            const contactArray = contactRaw ? contactRaw.split(CONTACT_SEPARATOR).map(s => s.trim()).filter(s => s !== '') : [];
            const newData = {
                isCommunicated: document.getElementById('modalIsCommunicated').value,
                category: document.getElementById('modalCategory').value.trim(),
                contactPerson: contactArray,
                feedback: document.getElementById('modalFeedback').value.trim(),
                dateRemark: document.getElementById('modalDateRemark').value.trim()
            };
            saveCategory(currentRefundNo, newData);
            try {
                if (currentImageFiles.length > 0) {
                    await saveImagesToDB(currentRefundNo, currentImageFiles);
                } else {
                    await deleteImagesFromDB(currentRefundNo);
                }
            } catch (e) {
                console.error('保存图片失败', e);
                alert('图片保存失败，请重试');
                return;
            }
            refreshRowCategoryCell(currentRefundNo);
            refreshFilterOptions();
            hideModal();
        };

        modalDiv.querySelector('#modalSaveBtn').addEventListener('click', saveHandler);
        modalDiv.querySelector('#modalCancelBtn').addEventListener('click', hideModal);
        modalDiv.addEventListener('click', (e) => { if (e.target === modalDiv) hideModal(); });

        modalInstance = { show: showModal, hide: hideModal };
        return modalInstance;
    }

    function openEditModal(refundNo) {
        createModal().show(refundNo);
    }

    function setupEventDelegate() {
        const container = document.querySelector('.exchange_tab');
        if (!container || container.hasAttribute('data-category-delegate')) return;
        container.setAttribute('data-category-delegate', 'true');
        container.addEventListener('click', (e) => {
            let btn = e.target;
            if (!btn.classList || !btn.classList.contains('edit-category-btn')) {
                btn = btn.closest('.edit-category-btn');
                if (!btn) return;
            }
            e.preventDefault();
            e.stopPropagation();
            const refundNo = btn.getAttribute('data-refundno') || getRefundNoFromRow(btn.closest('tr'));
            if (refundNo) openEditModal(refundNo);
            else alert('无法获取退换单号');
        });
    }

    function addCategoryColumn() {
        if (columnAdded) return;
        const headerTable = document.querySelector('.header_list table');
        const dataTable = document.querySelector('.inner_list table');
        if (!headerTable || !dataTable) return;
        const headerRow = headerTable.querySelector('tr.tab_bt') || headerTable.querySelector('tr');
        if (!headerRow) return;
        if (Array.from(headerRow.cells).some(cell => cell.textContent.trim() === '备注')) {
            columnAdded = true;
            setTimeout(() => refreshAllRows(), 100);
            return;
        }
        let insertIndex = headerRow.cells.length;
        for (let i = 0; i < headerRow.cells.length; i++) {
            if (headerRow.cells[i].textContent.trim() === ANCHOR_COLUMN_TEXT) {
                insertIndex = i + 1;
                break;
            }
        }
        const newHeaderCell = document.createElement('td');
        newHeaderCell.textContent = '备注';
        newHeaderCell.className = 'row05';
        newHeaderCell.style.width = DEFAULT_CATEGORY_WIDTH;
        newHeaderCell.style.backgroundColor = '#f5f7fa';
        newHeaderCell.style.textAlign = 'center';
        if (insertIndex >= headerRow.cells.length) headerRow.appendChild(newHeaderCell);
        else headerRow.insertBefore(newHeaderCell, headerRow.cells[insertIndex]);

        const rows = Array.from(dataTable.querySelectorAll('tbody tr'));
        const currentWidth = getCurrentCategoryColumnWidth();
        rows.forEach(row => {
            const refundNo = getRefundNoFromRow(row);
            const newTd = buildCategoryCell(refundNo, currentWidth);
            if (insertIndex >= row.cells.length) row.appendChild(newTd);
            else row.insertBefore(newTd, row.cells[insertIndex]);
        });

        syncTableColumnWidths();
        columnAdded = true;
    }

    function syncRowsWithCategory() {
        const dataTable = document.querySelector('.inner_list table');
        if (!dataTable) return;
        const headerTable = document.querySelector('.header_list table');
        if (!headerTable) return;
        const headerRow = headerTable.querySelector('tr.tab_bt') || headerTable.querySelector('tr');
        let categoryColIndex = -1;
        for (let i = 0; i < headerRow.cells.length; i++) {
            if (headerRow.cells[i].textContent.trim() === '备注') {
                categoryColIndex = i;
                break;
            }
        }
        if (categoryColIndex === -1) {
            if (!columnAdded) addCategoryColumn();
            return;
        }
        const rows = dataTable.querySelectorAll('tbody tr');
        const currentWidth = getCurrentCategoryColumnWidth();
        rows.forEach(row => {
            if (row.cells.length <= categoryColIndex || !row.cells[categoryColIndex].querySelector('.edit-category-btn')) {
                const refundNo = getRefundNoFromRow(row);
                if (!refundNo) return;
                const newTd = buildCategoryCell(refundNo, currentWidth);
                if (categoryColIndex >= row.cells.length) row.appendChild(newTd);
                else row.insertBefore(newTd, row.cells[categoryColIndex]);
            }
        });
    }

    function exportAllCachedData() {
        const allKeys = [];
        for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            if (key && key.startsWith(STORAGE_PREFIX)) allKeys.push(key);
        }
        if (allKeys.length === 0) { alert('没有找到任何已保存的备注记录。'); return; }
        const exportRows = [];
        for (const key of allKeys) {
            const refundNo = key.slice(STORAGE_PREFIX.length);
            const data = loadCategory(refundNo);
            const isEmpty = (!data.category || data.category.trim() === '') &&
                            data.contactPerson.length === 0 &&
                            (!data.feedback || data.feedback.trim() === '') &&
                            (data.isCommunicated === '—'|| !data.isCommunicated) &&
                            (!data.dateRemark || data.dateRemark.trim() === '');
            if (!isEmpty) {
                exportRows.push({
                    refundNo,
                    isCommunicated: data.isCommunicated,
                    category: data.category,
                    contactPerson: data.contactPerson.join('、'),
                    feedback: data.feedback,
                    dateRemark: data.dateRemark || ''
                });
            }
        }
        if (exportRows.length === 0) { alert('所有已保存的备注内容均为空，无可导出数据。'); return; }
        exportRows.sort((a, b) => (a.refundNo < b.refundNo ? 1 : (a.refundNo > b.refundNo ? -1 : 0)));
        const sheetData = [['退换单号', '沟通状态', '问题分类', '沟通人', '沟通反馈', '备注日期']];
        exportRows.forEach(r => sheetData.push([r.refundNo, r.isCommunicated, r.category, r.contactPerson, r.feedback, r.dateRemark]));
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '退换单备注');
        const now = new Date();
        const timeStr = `${now.getFullYear()}-${now.getMonth()+1}-${now.getDate()}_${now.getHours()}-${now.getMinutes()}-${now.getSeconds()}`;
        XLSX.writeFile(wb, `退换单备注_${timeStr}.xlsx`);
    }

    function addExportButton() {
        if (document.getElementById('custom_export_all_btn')) return;
        const targetLink = document.querySelector("span.gbtn_btn.le a[href^='javascript:bind_check_out_accountRepair_event']");
        let container = null;
        if (targetLink) container = targetLink.closest('span.gbtn_btn.le');
        if (container && container.parentElement) {
            const btnSpan = document.createElement('span');
            btnSpan.id = 'custom_export_all_btn';
            btnSpan.className = 'gbtn_btn le tm-green';
            btnSpan.style.marginLeft = '10px';
            const btnLink = document.createElement('a');
            btnLink.href = 'javascript:void(0)';
            btnLink.textContent = '导出备注';
            btnLink.style.cursor = 'pointer';
            btnLink.onclick = (e) => { e.preventDefault(); exportAllCachedData(); };
            btnSpan.appendChild(btnLink);
            container.parentElement.insertBefore(btnSpan, container.nextSibling);
            return;
        }
        const tabBar = document.querySelector('.exchange_tab .tab_bar');
        if (tabBar) {
            const btnDiv = document.createElement('div');
            btnDiv.id = 'custom_export_all_btn';
            btnDiv.style.display = 'inline-block';
            btnDiv.style.marginLeft = '15px';
            const btn = document.createElement('button');
            btn.textContent = '导出备注';
            btn.style.padding = '4px 12px';
            btn.style.cursor = 'pointer';
            btn.style.border = '1px solid #409eff';
            btn.style.background = '#409eff';
            btn.style.color = '#fff';
            btn.style.borderRadius = '4px';
            btn.onclick = exportAllCachedData;
            btnDiv.appendChild(btn);
            tabBar.appendChild(btnDiv);
            return;
        }
        const floatDiv = document.createElement('div');
        floatDiv.id = 'custom_export_all_btn';
        floatDiv.style.position = 'fixed';
        floatDiv.style.bottom = '20px';
        floatDiv.style.right = '20px';
        floatDiv.style.zIndex = '9999';
        const floatBtn = document.createElement('button');
        floatBtn.textContent = '导出备注';
        floatBtn.style.padding = '8px 16px';
        floatBtn.style.background = '#409eff';
        floatBtn.style.border = 'none';
        floatBtn.style.borderRadius = '6px';
        floatBtn.style.color = '#fff';
        floatBtn.style.cursor = 'pointer';
        floatBtn.style.boxShadow = '0 2px 6px rgba(0,0,0,0.2)';
        floatBtn.onclick = exportAllCachedData;
        floatDiv.appendChild(floatBtn);
        document.body.appendChild(floatDiv);
    }

    function observeTableChanges() {
        const targetNode = document.querySelector('.exchange_tab');
        if (!targetNode) return;
        const observer = new MutationObserver(() => {
            if (!columnAdded) {
                addCategoryColumn();
                addExportButton();
                addFilterControl();
                setupEventDelegate();
            } else {
                syncRowsWithCategory();
                refreshFilterOptions();
            }
        });
        observer.observe(targetNode, { childList: true, subtree: true });
    }

    function init() {
        waitForElement('.exchange_tab', () => {
            setTimeout(() => {
                addCategoryColumn();
                addExportButton();
                addFilterControl();
                setupEventDelegate();
                observeTableChanges();
                setTimeout(() => {
                    refreshAllRows();
                    refreshFilterOptions();
                }, 300);
            }, 200);
        });
    }

    function waitForElement(selector, callback, maxWait = 4000) {
        const start = Date.now();
        const timer = setInterval(() => {
            const el = document.querySelector(selector);
            if (el) { clearInterval(timer); callback(el); }
            else if (Date.now() - start > maxWait) { clearInterval(timer); console.warn('等待超时:', selector); }
        }, 150);
    }

    init();
})();

// 已点击问题链接的样式记录（保持不变）
(function() {
    'use strict';

    const STORAGE_KEY = 'ClickedQuestionInfoLinks';
    const CLICKED_CLASS = 'question-link-clicked';

    const style = document.createElement('style');
    style.textContent = `
        .${CLICKED_CLASS} {
            color: #333333 !important;
            padding-left: 4px !important;
        }
    `;
    document.head.appendChild(style);

    function getLinkIdentifier(anchor) {
        const onclickAttr = anchor.getAttribute('onclick');
        if (!onclickAttr) return null;
        const match = onclickAttr.match(/questionInfo\s*\(\s*['"]([^'"]+)['"]\s*,\s*['"]([^'"]+)['"]\s*\)/);
        if (match && match.length >= 3) {
            return `${match[1]}|${match[2]}`;
        }
        return null;
    }

    function getClickedSet() {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            try {
                return new Set(JSON.parse(stored));
            } catch(e) {
                return new Set();
            }
        }
        return new Set();
    }

    function saveClickedIdentifier(identifier) {
        if (!identifier) return;
        const clickedSet = getClickedSet();
        if (!clickedSet.has(identifier)) {
            clickedSet.add(identifier);
            localStorage.setItem(STORAGE_KEY, JSON.stringify([...clickedSet]));
        }
    }

    function applyClickedStyle(anchor, identifier) {
        if (!anchor) return;
        const clickedSet = getClickedSet();
        if (identifier && clickedSet.has(identifier)) {
            anchor.classList.add(CLICKED_CLASS);
        } else if (anchor.classList.contains(CLICKED_CLASS)) {
            anchor.classList.remove(CLICKED_CLASS);
        }
    }

    function markLinkAsClicked(anchor, identifier) {
        if (!anchor || !identifier) return false;
        if (anchor.classList.contains(CLICKED_CLASS)) return false;

        const clickedSet = getClickedSet();
        if (clickedSet.has(identifier)) {
            anchor.classList.add(CLICKED_CLASS);
            return false;
        }

        anchor.classList.add(CLICKED_CLASS);
        saveClickedIdentifier(identifier);
        return true;
    }

    function handleClick(event) {
        const anchor = event.target.closest('a');
        if (!anchor) return;

        const onclickAttr = anchor.getAttribute('onclick');
        if (!onclickAttr || !onclickAttr.includes('questionInfo')) return;

        const identifier = getLinkIdentifier(anchor);
        if (!identifier) return;

        if (anchor.hasAttribute('data-click-processed')) return;
        anchor.setAttribute('data-click-processed', 'true');

        setTimeout(() => {
            markLinkAsClicked(anchor, identifier);
        }, 10);
    }

    function initExistingLinks() {
        const anchors = document.querySelectorAll('a');
        anchors.forEach(anchor => {
            const onclickAttr = anchor.getAttribute('onclick');
            if (onclickAttr && onclickAttr.includes('questionInfo')) {
                const identifier = getLinkIdentifier(anchor);
                if (identifier) {
                    applyClickedStyle(anchor, identifier);
                }
            }
        });
    }

    function observeDynamicNodes() {
        const observer = new MutationObserver(mutations => {
            mutations.forEach(mutation => {
                mutation.addedNodes.forEach(node => {
                    if (node.nodeType === Node.ELEMENT_NODE) {
                        if (node.matches && node.matches('a')) {
                            const onclickAttr = node.getAttribute('onclick');
                            if (onclickAttr && onclickAttr.includes('questionInfo')) {
                                const identifier = getLinkIdentifier(node);
                                if (identifier) applyClickedStyle(node, identifier);
                            }
                        }
                        if (node.querySelectorAll) {
                            const links = node.querySelectorAll('a');
                            links.forEach(link => {
                                const onclickAttrLink = link.getAttribute('onclick');
                                if (onclickAttrLink && onclickAttrLink.includes('questionInfo')) {
                                    const identifier = getLinkIdentifier(link);
                                    if (identifier) applyClickedStyle(link, identifier);
                                }
                            });
                        }
                    }
                });
            });
        });
        if (document.body) {
            observer.observe(document.body, { childList: true, subtree: true });
        } else {
            window.addEventListener('DOMContentLoaded', () => {
                observer.observe(document.body, { childList: true, subtree: true });
            });
        }
    }

    document.addEventListener('click', handleClick, true);

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            initExistingLinks();
            observeDynamicNodes();
        });
    } else {
        initExistingLinks();
        observeDynamicNodes();
    }
})();
