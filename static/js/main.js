/**
 * 快手订单对账系统 - 前端交互逻辑
 */

// 全局状态
let state = {
    officialFile: null,
    customerFile: null,
    resultData: [],
    summary: {},
    currentPage: 1,
    pageSize: 50,
    filteredData: [],
    officialColumns: [],
    customerColumns: []
};

// DOM 元素
const elements = {
    officialFile: document.getElementById('official-file'),
    customerFile: document.getElementById('customer-file'),
    officialInfo: document.getElementById('official-info'),
    customerInfo: document.getElementById('customer-info'),
    officialColumns: document.getElementById('official-columns'),
    customerColumns: document.getElementById('customer-columns'),
    compareBtn: document.getElementById('compare-btn'),
    exportBtn: document.getElementById('export-btn'),
    progressArea: document.getElementById('progress-area'),
    errorArea: document.getElementById('error-area'),
    errorText: document.querySelector('.error-text'),
    summarySection: document.getElementById('summary-section'),
    resultSection: document.getElementById('result-section'),
    resultTbody: document.getElementById('result-tbody'),
    pagination: document.getElementById('pagination'),
    filterStatus: document.getElementById('filter-status'),
    searchInput: document.getElementById('search-input'),
    searchBtn: document.getElementById('search-btn')
};

/**
 * 初始化
 */
document.addEventListener('DOMContentLoaded', function() {
    bindEvents();
});

/**
 * 绑定事件
 */
function bindEvents() {
    // 文件上传
    elements.officialFile.addEventListener('change', (e) => handleFileUpload(e, 'official'));
    elements.customerFile.addEventListener('change', (e) => handleFileUpload(e, 'customer'));
    
    // 按钮事件
    elements.compareBtn.addEventListener('click', handleCompare);
    elements.exportBtn.addEventListener('click', handleExport);
    
    // 筛选和搜索
    elements.filterStatus.addEventListener('change', handleFilter);
    elements.searchBtn.addEventListener('click', handleSearch);
    elements.searchInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleSearch();
    });
}

/**
 * 处理文件上传
 */
async function handleFileUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;
    
    // 隐藏之前的错误
    hideError();
    showProgress('正在上传文件...');
    
    const formData = new FormData();
    formData.append('file', file);
    formData.append('type', type);
    
    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            // 保存文件信息
            if (type === 'official') {
                state.officialFile = data;
                showFileInfo('official', file.name, data.row_count);
                showColumnMapping('official', data.columns, data.detected);
            } else {
                state.customerFile = data;
                showFileInfo('customer', file.name, data.row_count);
                showColumnMapping('customer', data.columns, data.detected);
            }
            
            updateCompareButton();
        } else {
            showError(data.error);
            event.target.value = '';
        }
    } catch (error) {
        showError('上传失败: ' + error.message);
        event.target.value = '';
    } finally {
        hideProgress();
    }
}

/**
 * 显示文件信息
 */
function showFileInfo(type, filename, rowCount) {
    const infoEl = type === 'official' ? elements.officialInfo : elements.customerInfo;
    infoEl.querySelector('.filename').textContent = filename;
    infoEl.querySelector('.row-count').textContent = `${rowCount} 行数据`;
    infoEl.classList.remove('hidden');
}

/**
 * 显示列映射配置
 */
function showColumnMapping(type, columns, detected) {
    const container = type === 'official' ? elements.officialColumns : elements.customerColumns;
    
    // 填充下拉框
    const orderSelect = document.getElementById(`${type}-order`);
    const statusSelect = document.getElementById(`${type}-status`);
    const productSelect = document.getElementById(`${type}-product`);
    const amountSelect = document.getElementById(`${type}-amount`);
    const costSelect = document.getElementById(`${type}-cost`);
    
    // 清空并填充选项
    [orderSelect, productSelect, amountSelect].forEach(select => {
        select.innerHTML = columns.map(col => 
            `<option value="${col}">${col}</option>`
        ).join('');
    });
    
    if (type === 'official' && statusSelect) {
        statusSelect.innerHTML = columns.map(col => 
            `<option value="${col}">${col}</option>`
        ).join('');
    }
    
    if (costSelect) {
        costSelect.innerHTML = '<option value="">-- 请选择 --</option>' + 
            columns.map(col => `<option value="${col}">${col}</option>`).join('');
    }
    
    // 设置自动检测的默认值
    if (orderSelect && detected.order_id) orderSelect.value = detected.order_id;
    if (statusSelect && detected.status) statusSelect.value = detected.status;
    if (productSelect && detected.product) productSelect.value = detected.product;
    if (amountSelect && detected.amount) amountSelect.value = detected.amount;
    if (costSelect && detected.cost) costSelect.value = detected.cost;
    
    // 显示配置区域
    container.classList.remove('hidden');
}

/**
 * 获取列映射配置
 */
function getColumnMapping(type) {
    const mapping = {
        order_id: document.getElementById(`${type}-order`).value,
        product: document.getElementById(`${type}-product`).value,
        amount: document.getElementById(`${type}-amount`).value,
        cost: document.getElementById(`${type}-cost`)?.value || null
    };
    
    if (type === 'official') {
        mapping.status = document.getElementById(`${type}-status`).value;
    }
    
    return mapping;
}

/**
 * 更新对比按钮状态
 */
function updateCompareButton() {
    const canCompare = state.officialFile && state.customerFile;
    elements.compareBtn.disabled = !canCompare;
}

/**
 * 处理对比
 */
async function handleCompare() {
    hideError();
    showProgress('正在对比数据，请稍候...');
    
    // 获取列映射
    const officialMapping = getColumnMapping('official');
    const customerMapping = getColumnMapping('customer');
    
    // 验证必填项
    if (!officialMapping.order_id || !officialMapping.status || !officialMapping.product || !officialMapping.amount) {
        showError('请完整配置官方表的列映射');
        hideProgress();
        return;
    }
    
    if (!customerMapping.order_id || !customerMapping.product || !customerMapping.amount || !customerMapping.cost) {
        showError('请完整配置客服表的列映射（成本列为必填）');
        hideProgress();
        return;
    }
    
    try {
        const response = await fetch('/api/compare', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                official_file: state.officialFile.filepath,
                customer_file: state.customerFile.filepath,
                official_mapping: officialMapping,
                customer_mapping: customerMapping
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            state.resultData = data.data;
            state.summary = data.summary;
            state.filteredData = [...state.resultData];
            state.currentPage = 1;
            
            // 显示结果
            showSummary(data.summary);
            showResultTable();
            elements.exportBtn.disabled = false;
            
            // 如果有重复订单，显示警告
            if (data.duplicates && data.duplicates.length > 0) {
                showError(`检测到 ${data.duplicates.length} 个重复订单号，已自动去重`);
            }
        } else {
            showError(data.error);
        }
    } catch (error) {
        showError('对比失败: ' + error.message);
    } finally {
        hideProgress();
    }
}

/**
 * 显示汇总统计
 */
function showSummary(summary) {
    document.getElementById('summary-sales').textContent = formatMoney(summary.total_sales);
    document.getElementById('summary-cost').textContent = formatMoney(summary.total_cost);
    document.getElementById('summary-profit').textContent = formatMoney(summary.total_profit);
    document.getElementById('summary-total').textContent = summary.total_orders;
    document.getElementById('summary-normal').textContent = summary.normal_orders;
    document.getElementById('summary-abnormal').textContent = summary.abnormal_orders;
    document.getElementById('summary-missing').textContent = summary.missing_orders;
    document.getElementById('summary-extra').textContent = summary.extra_orders;
    
    elements.summarySection.classList.remove('hidden');
}

/**
 * 显示结果表格
 */
function showResultTable() {
    renderTable();
    elements.resultSection.classList.remove('hidden');
    
    // 滚动到结果区域
    elements.resultSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * 渲染表格
 */
function renderTable() {
    const start = (state.currentPage - 1) * state.pageSize;
    const end = start + state.pageSize;
    const pageData = state.filteredData.slice(start, end);
    
    elements.resultTbody.innerHTML = pageData.map((row, index) => {
        const isLoss = row.利润 < 0;
        const isAbnormal = row.状态 === '异常';
        const globalIndex = start + index + 1;
        
        return `
            <tr class="${isAbnormal ? 'abnormal' : ''}">
                <td>${globalIndex}</td>
                <td>${row.订单号}</td>
                <td style="text-align: left;">${row.商品名称}</td>
                <td>${formatMoney(row.销售额)}</td>
                <td>${formatMoney(row.成本)}</td>
                <td class="${isLoss ? 'loss' : ''}">${formatMoney(row.利润)}</td>
                <td class="${isAbnormal ? 'status-abnormal' : 'status-normal'}">${row.状态}</td>
                <td>${row.备注}</td>
            </tr>
        `;
    }).join('');
    
    renderPagination();
}

/**
 * 渲染分页
 */
function renderPagination() {
    const totalPages = Math.ceil(state.filteredData.length / state.pageSize);
    
    if (totalPages <= 1) {
        elements.pagination.innerHTML = '';
        return;
    }
    
    let html = '';
    
    // 上一页
    html += `<button ${state.currentPage === 1 ? 'disabled' : ''} onclick="changePage(${state.currentPage - 1})">上一页</button>`;
    
    // 页码
    html += `<span class="page-info">第 ${state.currentPage} / ${totalPages} 页，共 ${state.filteredData.length} 条</span>`;
    
    // 下一页
    html += `<button ${state.currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${state.currentPage + 1})">下一页</button>`;
    
    elements.pagination.innerHTML = html;
}

/**
 * 切换页面
 */
function changePage(page) {
    const totalPages = Math.ceil(state.filteredData.length / state.pageSize);
    if (page < 1 || page > totalPages) return;
    
    state.currentPage = page;
    renderTable();
    
    // 滚动到表格顶部
    document.querySelector('.table-container').scrollIntoView({ behavior: 'smooth' });
}

/**
 * 处理筛选
 */
function handleFilter() {
    const status = elements.filterStatus.value;
    
    if (status === 'all') {
        state.filteredData = [...state.resultData];
    } else if (status === 'normal') {
        state.filteredData = state.resultData.filter(row => row.状态 === '正常');
    } else if (status === 'abnormal') {
        state.filteredData = state.resultData.filter(row => row.状态 === '异常');
    }
    
    state.currentPage = 1;
    renderTable();
}

/**
 * 处理搜索
 */
function handleSearch() {
    const keyword = elements.searchInput.value.trim().toLowerCase();
    
    if (!keyword) {
        state.filteredData = [...state.resultData];
    } else {
        state.filteredData = state.resultData.filter(row => 
            row.订单号.toLowerCase().includes(keyword)
        );
    }
    
    state.currentPage = 1;
    renderTable();
}

/**
 * 处理导出
 */
async function handleExport() {
    if (state.resultData.length === 0) {
        showError('没有数据可导出');
        return;
    }
    
    showProgress('正在生成Excel文件...');
    
    try {
        const response = await fetch('/api/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                data: state.resultData,
                summary: state.summary
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            // 触发下载
            window.location.href = data.download_url;
        } else {
            showError(data.error);
        }
    } catch (error) {
        showError('导出失败: ' + error.message);
    } finally {
        hideProgress();
    }
}

/**
 * 显示进度
 */
function showProgress(text) {
    document.querySelector('.progress-text').textContent = text;
    elements.progressArea.classList.remove('hidden');
}

/**
 * 隐藏进度
 */
function hideProgress() {
    elements.progressArea.classList.add('hidden');
}

/**
 * 显示错误
 */
function showError(message) {
    elements.errorText.textContent = message;
    elements.errorArea.classList.remove('hidden');
}

/**
 * 隐藏错误
 */
function hideError() {
    elements.errorArea.classList.add('hidden');
}

/**
 * 格式化金额
 */
function formatMoney(amount) {
    if (amount === undefined || amount === null) return '¥0.00';
    return '¥' + Number(amount).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}
