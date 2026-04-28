/**********************************************
 * 全局字段列表
 **********************************************/
const FIELDS = [
  '品牌名称', '产品名称', '产品型号', '产品价格',
  '报价日期', '报价人', '联系方式', '所属部门',
  '负责区域', '客户名称', '关键参数', '配置文档路径'
];

/**********************************************
 * IndexedDB 封装
 **********************************************/
const DB_NAME = 'ProductQuoteDB';
const DB_VERSION = 1;
const STORE_NAME = 'products';
let db = null;

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        const store = db.createObjectStore(STORE_NAME, {
          keyPath: 'id',
          autoIncrement: true
        });
        store.createIndex('品牌名称', '品牌名称', { unique: false });
        store.createIndex('产品名称', '产品名称', { unique: false });
        store.createIndex('产品型号', '产品型号', { unique: false });
      }
    };
    request.onsuccess = (e) => {
      db = e.target.result;
      resolve(db);
    };
    request.onerror = (e) => reject(e.target.error);
  });
}

function getAllProducts() {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result);
    request.onerror = (e) => reject(e.target.error);
  });
}

function addProducts(products) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    for (const p of products) {
      store.add(p);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

function deleteProduct(id) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const request = store.delete(id);
    request.onsuccess = () => resolve();
    request.onerror = (e) => reject(e.target.error);
  });
}

function deleteMultipleProducts(ids) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    for (const id of ids) {
      store.delete(id);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = (e) => reject(e.target.error);
  });
}

/**********************************************
 * Excel 日期转换 (1900 日期系统)
 **********************************************/
function excelSerialToDate(serial) {
  const date = new Date((serial - 25569) * 86400000);
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function normalizeDate(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number' && value > 30000) {
    return excelSerialToDate(value);
  }
  return String(value).split(' ')[0];
}

/**********************************************
 * 全局状态
 **********************************************/
let allProducts = [];
let filteredProducts = [];
let duplicateGroupIndices = new Set();

/**********************************************
 * DOM 元素
 **********************************************/
const searchInput = document.getElementById('searchInput');
const searchBtn = document.getElementById('searchBtn');
const clearSearchBtn = document.getElementById('clearSearchBtn');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const checkDuplicateBtn = document.getElementById('checkDuplicateBtn');
const batchDeleteBtn = document.getElementById('batchDeleteBtn');
const tableHeaderRow = document.getElementById('tableHeaderRow');
const tableBody = document.getElementById('tableBody');
const noDataDiv = document.getElementById('noData');
const addForm = document.getElementById('addRecordForm');
const sortBtns = document.querySelectorAll('.sort-btn');
const colChecks = document.querySelectorAll('.col-check');

/**********************************************
 * 工具函数
 **********************************************/
function getVisibleColumns() {
  const checked = [];
  colChecks.forEach(cb => {
    if (cb.checked) checked.push(cb.value);
  });
  return checked;
}

function filterBySearch(products, keyword) {
  if (!keyword.trim()) return products;
  const kw = keyword.trim().toLowerCase();
  return products.filter(p => {
    const brand = String(p['品牌名称'] || '').toLowerCase();
    const name = String(p['产品名称'] || '').toLowerCase();
    const model = String(p['产品型号'] || '').toLowerCase();
    return brand.includes(kw) || name.includes(kw) || model.includes(kw);
  });
}

function sortProducts(products, field) {
  return [...products].sort((a, b) => {
    let valA = a[field] || '';
    let valB = b[field] || '';
    if (field === '产品价格') {
      return parseFloat(valA) - parseFloat(valB);
    } else if (field === '报价日期') {
      return new Date(valA) - new Date(valB);
    } else {
      return String(valA).localeCompare(String(valB), 'zh-CN');
    }
  });
}

function isSameRow(row1, row2) {
  for (const field of FIELDS) {
    if (String(row1[field] || '') !== String(row2[field] || '')) return false;
  }
  return true;
}

function findDuplicateIds(products) {
  const seen = [];
  const duplicateIds = new Set();
  for (const p of products) {
    const firstIndex = seen.findIndex(s => isSameRow(s, p));
    if (firstIndex !== -1) {
      duplicateIds.add(seen[firstIndex].id);
      duplicateIds.add(p.id);
    } else {
      seen.push(p);
    }
  }
  return duplicateIds;
}

/**********************************************
 * 渲染表格（含价格格式化）
 **********************************************/
function renderTable(products) {
  const columns = getVisibleColumns();
  tableHeaderRow.innerHTML = '';

  // 复选框列
  const thCheck = document.createElement('th');
  thCheck.className = 'check-col';
  const allCheck = document.createElement('input');
  allCheck.type = 'checkbox';
  allCheck.id = 'selectAllCheckbox';
  thCheck.appendChild(allCheck);
  tableHeaderRow.appendChild(thCheck);

  // 数据列
  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    tableHeaderRow.appendChild(th);
  });

  const thAction = document.createElement('th');
  thAction.textContent = '操作';
  tableHeaderRow.appendChild(thAction);

  tableBody.innerHTML = '';
  if (products.length === 0) {
    noDataDiv.style.display = 'block';
    return;
  }
  noDataDiv.style.display = 'none';

  products.forEach(product => {
    const tr = document.createElement('tr');
    if (duplicateGroupIndices.has(product.id)) {
      tr.classList.add('duplicate');
    }

    // 复选框
    const tdCheck = document.createElement('td');
    tdCheck.className = 'check-col';
    const rowCheck = document.createElement('input');
    rowCheck.type = 'checkbox';
    rowCheck.className = 'row-checkbox';
    rowCheck.dataset.id = product.id;
    tdCheck.appendChild(rowCheck);
    tr.appendChild(tdCheck);

    // 数据列
    columns.forEach(col => {
      const td = document.createElement('td');
      let cellValue = product[col] || '';
      // 价格列特殊格式化
      if (col === '产品价格' && cellValue !== '') {
        const num = parseFloat(cellValue);
        if (!isNaN(num)) {
          cellValue = num.toLocaleString('zh-CN', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
          });
        }
      }
      td.textContent = cellValue;
      tr.appendChild(td);
    });

    // 删除按钮
    const tdAction = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.className = 'delete-row-btn';
    delBtn.textContent = '删除';
    delBtn.addEventListener('click', async () => {
      if (confirm('确定要删除这条记录吗？')) {
        try {
          await deleteProduct(product.id);
          await refreshData();
        } catch (err) {
          alert('删除失败：' + err.message);
        }
      }
    });
    tdAction.appendChild(delBtn);
    tr.appendChild(tdAction);
    tableBody.appendChild(tr);
  });

  // 全选逻辑
  allCheck.addEventListener('change', () => {
    const rowChecks = document.querySelectorAll('.row-checkbox');
    rowChecks.forEach(cb => cb.checked = allCheck.checked);
  });
}

/**********************************************
 * 刷新数据
 **********************************************/
async function refreshData() {
  try {
    allProducts = await getAllProducts();
    const keyword = searchInput.value;
    let result = filterBySearch(allProducts, keyword);
    const activeSort = document.querySelector('.sort-btn.active');
    if (activeSort) {
      result = sortProducts(result, activeSort.dataset.sort);
    }
    filteredProducts = result;
    duplicateGroupIndices = findDuplicateIds(filteredProducts);
    renderTable(filteredProducts);
  } catch (err) {
    console.error('刷新数据失败：', err);
    alert('数据加载失败，请检查控制台');
  }
}

/**********************************************
 * 事件绑定
 **********************************************/
searchBtn.addEventListener('click', refreshData);
clearSearchBtn.addEventListener('click', () => {
  searchInput.value = '';
  refreshData();
});
searchInput.addEventListener('keypress', (e) => {
  if (e.key === 'Enter') refreshData();
});

colChecks.forEach(cb => {
  cb.addEventListener('change', () => renderTable(filteredProducts));
});

sortBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    sortBtns.forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filteredProducts = sortProducts(filteredProducts, btn.dataset.sort);
    renderTable(filteredProducts);
  });
});

// ---------- 上传 Excel ----------
uploadBtn.addEventListener('click', async () => {
  const file = fileInput.files[0];
  if (!file) {
    alert('请先选择一个Excel文件');
    return;
  }
  try {
    const data = await readExcel(file);
    console.log('读取原始数据样本（前2条）:', data.slice(0, 2));
    if (data.length === 0) {
      alert('Excel没有数据');
      return;
    }
    const firstRow = data[0];
    const missing = FIELDS.filter(f => !(f in firstRow));
    if (missing.length > 0) {
      if (!confirm(`Excel中缺少以下字段：${missing.join('、')}\n是否仍然导入？`)) {
        return;
      }
    }
    const toAdd = data.map(row => {
      const obj = {};
      FIELDS.forEach(f => {
        let val = row[f];
        if (f === '报价日期') {
          val = normalizeDate(val);
        }
        obj[f] = val !== undefined && val !== null ? val : '';
      });
      return obj;
    });
    console.log('转换后待导入数据（前2条）:', toAdd.slice(0, 2));
    await addProducts(toAdd);
    fileInput.value = '';
    alert(`成功导入 ${toAdd.length} 条记录`);
    await refreshData();
  } catch (err) {
    console.error('导入失败：', err);
    alert('读取Excel失败：' + err.message);
  }
});

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const json = XLSX.utils.sheet_to_json(ws);
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ---------- 查重 ----------
checkDuplicateBtn.addEventListener('click', async () => {
  duplicateGroupIndices = findDuplicateIds(filteredProducts);
  if (duplicateGroupIndices.size === 0) {
    alert('✅ 没有发现重复数据！');
  } else {
    alert(`🔍 发现 ${duplicateGroupIndices.size} 条重复数据（已红色高亮），可勾选后批量删除。`);
  }
  renderTable(filteredProducts);
});

// ---------- 批量删除 ----------
batchDeleteBtn.addEventListener('click', async () => {
  const checkedBoxes = document.querySelectorAll('.row-checkbox:checked');
  if (checkedBoxes.length === 0) {
    alert('请先勾选要删除的行');
    return;
  }
  if (!confirm(`确定要删除选中的 ${checkedBoxes.length} 条记录吗？`)) return;
  const idsToDelete = Array.from(checkedBoxes).map(cb => Number(cb.dataset.id));
  try {
    await deleteMultipleProducts(idsToDelete);
    duplicateGroupIndices.clear();
    await refreshData();
    alert(`✅ 已成功删除 ${idsToDelete.length} 条记录`);
  } catch (err) {
    alert('批量删除失败：' + err.message);
  }
});

// ---------- 手动添加 ----------
addForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(addForm);
  const newRec = {};
  FIELDS.forEach(f => {
    let val = formData.get(f) || '';
    if (f === '报价日期') val = normalizeDate(val);
    newRec[f] = val;
  });
  try {
    await addProducts([newRec]);
    addForm.reset();
    alert('添加成功！');
    await refreshData();
  } catch (err) {
    alert('添加失败：' + err.message);
  }
});

/**********************************************
 * 启动
 **********************************************/
(async function init() {
  try {
    await openDB();
    console.log('数据库已就绪');
    await refreshData();
  } catch (err) {
    console.error('初始化失败', err);
    alert('数据库初始化失败，请检查浏览器是否支持IndexedDB');
  }
})();

// 注册 Service Worker（仅当通过 http/https 访问时）
if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/service-worker.js')
      .then(reg => console.log('✅ Service Worker 注册成功'))
      .catch(err => console.log('❌ Service Worker 注册失败:', err));
  });
} else if (location.protocol === 'file:') {
  console.log('ℹ️ 通过 file:// 协议打开，离线安装功能不可用，其他功能正常。');
}