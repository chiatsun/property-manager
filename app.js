const API_URL = "https://script.google.com/macros/s/AKfycbyJkPfvNA0gosQeGaN7INsRVe82P-fCkN4ZWSenHSviMh-6pUYv8IB4vEmStYiYwPLg1w/exec";

// 初始化
window.onload = function() {
  loadOptions();
};

// Tab Switching
function switchTab(tabId) {
  document.querySelectorAll('.tab-pane').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  
  document.getElementById(tabId).classList.add('active');
  const btnIndex = tabId === 'add-tab' ? 0 : tabId === 'query-tab' ? 1 : 2;
  document.querySelectorAll('.tab-btn')[btnIndex].classList.add('active');

  // Stop scanner if active
  if (html5QrCode) {
    html5QrCode.stop().catch(err => console.log("Stop scanner error", err));
    document.getElementById('reader').style.display = 'none';
  }
}

// Loading Spinner
function showLoading() {
  document.getElementById('loading-overlay').style.display = 'flex';
}

function hideLoading() {
  document.getElementById('loading-overlay').style.display = 'none';
}

// Image Preview & Compression
function previewImage(input, previewId) {
  const preview = document.getElementById(previewId);
  if (input.files && input.files[0]) {
    const reader = new FileReader();
    reader.onload = function(e) {
      preview.src = e.target.result;
      preview.style.display = 'block';
    }
    reader.readAsDataURL(input.files[0]);
  } else {
    preview.style.display = 'none';
  }
}

function compressImage(file) {
  return new Promise((resolve) => {
    if (!file) {
      resolve(null);
      return;
    }
    const reader = new FileReader();
    reader.onload = function(event) {
      const img = new Image();
      img.onload = function() {
        const canvas = document.createElement('canvas');
        const MAX_WIDTH = 1200;
        const MAX_HEIGHT = 1200;
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > MAX_WIDTH) {
            height *= MAX_WIDTH / width;
            width = MAX_WIDTH;
          }
        } else {
          if (height > MAX_HEIGHT) {
            width *= MAX_HEIGHT / height;
            height = MAX_HEIGHT;
          }
        }
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0, width, height);
        // Compress to JPEG, quality 0.7
        resolve(canvas.toDataURL('image/jpeg', 0.7));
      }
      img.src = event.target.result;
    }
    reader.readAsDataURL(file);
  });
}

// Form Submission
async function handleFormSubmit(e) {
  e.preventDefault();
  showLoading();

  const data = {
    category: document.getElementById('category').value,
    id: document.getElementById('id').value,
    name: document.getElementById('name').value,
    alias: document.getElementById('alias').value,
    brand: document.getElementById('brand').value,
    unit: document.getElementById('unit').value,
    acquireDate: document.getElementById('acquireDate').value,
    lifespan: document.getElementById('lifespan').value,
    location: document.getElementById('location').value,
    userDept: document.getElementById('userDept').value,
    notes: document.getElementById('notes').value,
    scrapDate: document.getElementById('scrapDate').value,
    isScrapped: document.getElementById('isScrapped').checked
  };

  const p1File = document.getElementById('photo1').files[0];
  const p2File = document.getElementById('photo2').files[0];

  try {
    const p1Base64 = await compressImage(p1File);
    const p2Base64 = await compressImage(p2File);

    fetch(API_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'saveProperty',
        data: data,
        p1Base64: p1Base64,
        p2Base64: p2Base64
      })
    })
    .then(res => res.json())
    .then(response => {
      hideLoading();
      if(response.success) {
        alert("資料新增成功！");
        document.getElementById('add-form').reset();
        document.getElementById('preview1').style.display = 'none';
        document.getElementById('preview2').style.display = 'none';
        loadOptions(); // 儲存成功後更新下拉選單
      } else {
        alert("新增失敗：" + response.message);
      }
    })
    .catch(err => {
      hideLoading();
      alert("連線錯誤：" + err);
    });
  } catch (err) {
    hideLoading();
    alert("圖片處理錯誤：" + err);
  }
}

// Barcode Scanning
let html5QrCode;
function startScan(targetInputId) {
  const readerElement = document.getElementById('reader');
  readerElement.style.display = 'block';

  if (!html5QrCode) {
    html5QrCode = new Html5Qrcode("reader");
  }

  const qrCodeSuccessCallback = (decodedText, decodedResult) => {
    document.getElementById(targetInputId).value = decodedText;
    html5QrCode.stop().then(() => {
      readerElement.style.display = 'none';
    }).catch(err => {
      console.log("Failed to stop scanner", err);
    });
  };

  const config = { fps: 10, qrbox: { width: 250, height: 250 } };
  html5QrCode.start({ facingMode: "environment" }, config, qrCodeSuccessCallback)
    .catch(err => {
      alert("無法啟動相機，請確認已授權權限：" + err);
      readerElement.style.display = 'none';
    });
}

// Search Logic
function handleSearchEnter(e) {
  if (e.key === 'Enter') {
    searchData();
  }
}

function searchData() {
  const query = document.getElementById('searchInput').value;
  showLoading();
  
  fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({
      action: 'searchProperties',
      query: query
    })
  })
  .then(res => res.json())
  .then(response => {
    if(response.success) {
      renderSearchResults(response.data);
    } else {
      hideLoading();
      alert("搜尋失敗：" + response.message);
    }
  })
  .catch(err => {
    hideLoading();
    alert("搜尋連線錯誤：" + err);
  });
}

let searchResultsCache = [];

function renderSearchResults(results) {
  hideLoading();
  searchResultsCache = results;
  const container = document.getElementById('results-container');
  
  if (results.length === 0) {
    container.innerHTML = '<div style="text-align:center; color:#6b7280;">找不到符合的資料</div>';
    return;
  }

  let html = '';
  results.forEach((item, index) => {
    let cardClass = 'data-card ';
    if (item.isScrapped) {
      cardClass += 'status-scrapped'; // 紅底
    } else if (item.category === '財產') {
      cardClass += 'type-property'; // 白底
    } else {
      cardClass += 'type-item'; // 黃底
    }

    html += `
      <div class="${cardClass}" onclick="openDetailModal(${index})">
        <div class="card-header">
          <span class="card-title">${item.name}</span>
          <span class="card-id">${item.id}</span>
        </div>
        <div class="card-body">
          類別：${item.category} | 狀態：${item.isScrapped ? '已報廢' : '使用中'}<br>
          地點：${item.location || '未設定'} | 報廢日：${item.scrapDate ? new Date(item.scrapDate).toLocaleDateString() : '無'}
        </div>
      </div>
    `;
  });
  
  container.innerHTML = html;
}

// Modal & Lightbox
function openDetailModal(index) {
  const item = searchResultsCache[index];
  if (!item) return;

  const body = document.getElementById('modal-body');
  
  const generateRow = (label, val) => `
    <div class="detail-row">
      <div class="detail-label">${label}</div>
      <div class="detail-value">${val || '-'}</div>
    </div>
  `;

  let html = `
    ${generateRow('編號', item.id)}
    ${generateRow('名稱', item.name)}
    ${generateRow('別名', item.alias)}
    ${generateRow('類別', item.category)}
    ${generateRow('型式/廠牌', item.brand)}
    ${generateRow('數量單位', item.unit)}
    ${generateRow('取得日期', item.acquireDate ? new Date(item.acquireDate).toLocaleDateString() : '')}
    ${generateRow('使用年限', item.lifespan ? item.lifespan + ' 年' : '')}
    ${generateRow('存置地點', item.location)}
    ${generateRow('使用人/單位', item.userDept)}
    ${generateRow('報廢日期', item.scrapDate ? new Date(item.scrapDate).toLocaleDateString() : '')}
    ${generateRow('報廢狀態', item.isScrapped ? '已報廢 (紅底)' : '未報廢')}
    ${generateRow('備註', item.notes)}
  `;

  let imagesHtml = '<div class="detail-images">';
  if (item.photo1) {
    imagesHtml += `<img src="${item.photo1}" class="detail-img" onclick="openLightbox('${item.photo1}')">`;
  }
  if (item.photo2) {
    imagesHtml += `<img src="${item.photo2}" class="detail-img" onclick="openLightbox('${item.photo2}')">`;
  }
  imagesHtml += '</div>';

  if (item.photo1 || item.photo2) {
    html += imagesHtml;
  }

  body.innerHTML = html;
  
  // 加入按鈕區
  const footer = document.createElement('div');
  footer.style.marginTop = '1.5rem';
  footer.style.display = 'flex';
  footer.style.gap = '10px';
  
  const editBtn = document.createElement('button');
  editBtn.className = 'btn-primary';
  editBtn.style.background = '#6b7280';
  editBtn.innerText = '修改資料';
  editBtn.onclick = () => checkPasswordAndEdit(index);
  
  footer.appendChild(editBtn);
  body.appendChild(footer);

  document.getElementById('detail-modal').classList.add('active');
}

// 密碼驗證與開啟編輯模式
async function checkPasswordAndEdit(index) {
  const password = prompt("請輸入系統密碼以進行修改：");
  if (password === null) return;
  
  showLoading();
  fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({ action: 'getAppConfig' })
  })
  .then(res => res.json())
  .then(response => {
    hideLoading();
    if (response.success && response.data.password === password) {
      renderEditMode(index);
    } else {
      alert("密碼錯誤，無法修改！");
    }
  })
  .catch(err => {
    hideLoading();
    alert("驗證失敗：" + err);
  });
}

function renderEditMode(index) {
  const item = searchResultsCache[index];
  const body = document.getElementById('modal-body');
  
  const generateEditRow = (label, id, type = 'text', val = '', options = null) => {
    let inputHtml = `<input type="${type}" id="edit-${id}" value="${val || ''}" style="width:100%; padding:5px;">`;
    if (type === 'select' && options) {
      inputHtml = `<select id="edit-${id}" style="width:100%; padding:5px;">
        ${options.map(opt => `<option value="${opt}" ${opt === val ? 'selected' : ''}>${opt}</option>`).join('')}
      </select>`;
    }
    if (type === 'checkbox') {
        inputHtml = `<input type="checkbox" id="edit-${id}" ${val ? 'checked' : ''}>`;
    }
    return `
      <div class="detail-row" style="flex-direction:column; align-items:flex-start;">
        <div class="detail-label">${label}</div>
        <div class="detail-value" style="width:100%;">${inputHtml}</div>
      </div>
    `;
  };

  let html = `
    ${generateEditRow('類別', 'category', 'select', item.category, ['財產', '物品'])}
    ${generateEditRow('編號', 'id', 'text', item.id)}
    ${generateEditRow('名稱', 'name', 'text', item.name)}
    ${generateEditRow('別名', 'alias', 'text', item.alias)}
    ${generateEditRow('型式/廠牌', 'brand', 'text', item.brand)}
    ${generateEditRow('數量單位', 'unit', 'text', item.unit)}
    ${generateEditRow('取得日期', 'acquireDate', 'date', item.acquireDate ? new Date(item.acquireDate).toISOString().split('T')[0] : '')}
    ${generateEditRow('使用年限', 'lifespan', 'number', item.lifespan)}
    ${generateEditRow('存置地點', 'location', 'text', item.location)}
    ${generateEditRow('使用人/單位', 'userDept', 'text', item.userDept)}
    ${generateEditRow('報廢日期', 'scrapDate', 'date', item.scrapDate ? new Date(item.scrapDate).toISOString().split('T')[0] : '')}
    ${generateEditRow('是否完成報廢', 'isScrapped', 'checkbox', item.isScrapped)}
    <div class="detail-row" style="flex-direction:column;">
      <div class="detail-label">更換照片 1 (不更換請留空)</div>
      <input type="file" id="edit-photo1" accept="image/*" capture="environment">
    </div>
    <div class="detail-row" style="flex-direction:column;">
      <div class="detail-label">更換照片 2 (不更換請留空)</div>
      <input type="file" id="edit-photo2" accept="image/*" capture="environment">
    </div>
  `;

  body.innerHTML = html;
  
  const footer = document.createElement('div');
  footer.style.marginTop = '1.5rem';
  footer.style.display = 'flex';
  footer.style.gap = '10px';
  
  const saveBtn = document.createElement('button');
  saveBtn.className = 'btn-primary';
  saveBtn.innerText = '儲存修改';
  saveBtn.onclick = () => handleUpdateSubmit(index);
  
  const cancelBtn = document.createElement('button');
  cancelBtn.className = 'btn-primary';
  cancelBtn.style.background = '#e5e7eb';
  cancelBtn.style.color = '#374151';
  cancelBtn.innerText = '取消';
  cancelBtn.onclick = () => openDetailModal(index);
  
  footer.appendChild(saveBtn);
  footer.appendChild(cancelBtn);
  body.appendChild(footer);
}

async function handleUpdateSubmit(index) {
  const item = searchResultsCache[index];
  showLoading();

  const data = {
    category: document.getElementById('edit-category').value,
    id: document.getElementById('edit-id').value,
    name: document.getElementById('edit-name').value,
    alias: document.getElementById('edit-alias').value,
    brand: document.getElementById('edit-brand').value,
    unit: document.getElementById('edit-unit').value,
    acquireDate: document.getElementById('edit-acquireDate').value,
    lifespan: document.getElementById('edit-lifespan').value,
    location: document.getElementById('edit-location').value,
    userDept: document.getElementById('edit-userDept').value,
    scrapDate: document.getElementById('edit-scrapDate').value,
    isScrapped: document.getElementById('edit-isScrapped').checked,
    photo1: item.photo1, // 預設保留舊連結
    photo2: item.photo2
  };

  const p1File = document.getElementById('edit-photo1').files[0];
  const p2File = document.getElementById('edit-photo2').files[0];

  try {
    const p1Base64 = p1File ? await compressImage(p1File) : null;
    const p2Base64 = p2File ? await compressImage(p2File) : null;

    fetch(API_URL, {
      method: 'POST',
      body: JSON.stringify({
        action: 'updateProperty',
        rowIdx: item.rowIdx,
        data: data,
        p1Base64: p1Base64,
        p2Base64: p2Base64
      })
    })
    .then(res => res.json())
    .then(response => {
      hideLoading();
      if(response.success) {
        alert("修改成功！");
        closeModal();
        searchData(); // 重新搜尋以更新畫面
      } else {
        alert("修改失敗：" + response.message);
      }
    })
    .catch(err => {
      hideLoading();
      alert("更新連線錯誤：" + err);
    });
  } catch (err) {
    hideLoading();
    alert("圖片處理錯誤：" + err);
  }
}

function closeModal() {
  document.getElementById('detail-modal').classList.remove('active');
}

function openLightbox(url) {
  document.getElementById('lightbox-img').src = url;
  document.getElementById('lightbox').classList.add('active');
}

function closeLightbox() {
  document.getElementById('lightbox').classList.remove('active');
}

// Scrap Report Generation
function generateScrapReport() {
  const date = document.getElementById('targetScrapDate').value;
  const category = document.getElementById('targetCategory').value;

  if (!date) {
    alert("請選擇報廢日期");
    return;
  }

  showLoading();
  fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({
      action: 'generateScrapReport',
      targetDate: date,
      targetCategory: category
    })
  })
  .then(res => res.json())
  .then(response => {
    hideLoading();
    const resultDiv = document.getElementById('scrap-result');
    if (response.success) {
      resultDiv.innerHTML = `
        <div style="color: #10b981; font-weight: bold; margin-bottom: 10px;">${response.message}</div>
        <a href="${response.sheetUrl}" target="_blank" class="btn-primary" style="display:inline-block; text-decoration:none;">前往下載/查看報廢單</a>
        <div style="margin-top:10px; font-size:0.9rem; color:#6b7280;">提示：點擊連結開啟 Google 試算表後，可透過「檔案 > 下載 > OpenDocument 格式 (.ods)」下載為 ODS 檔。</div>
      `;
    } else {
      resultDiv.innerHTML = `<div style="color: #ef4444;">${response.message}</div>`;
    }
  })
  .catch(err => {
    hideLoading();
    alert("產出報廢單失敗：" + err);
  });
}

// 載入存置地點與使用人的下拉選單
function loadOptions() {
  fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({ action: 'getOptions' })
  })
  .then(res => res.json())
  .then(response => {
    if (response.success && response.data) {
      const options = response.data;
      if (options.locations) {
        updateDataList('location-list', options.locations);
      }
      if (options.users) {
        updateDataList('user-list', options.users);
      }
    }
  })
  .catch(err => console.log("載入選單失敗", err));
}

function updateDataList(id, options) {
  const list = document.getElementById(id);
  if (!list) return;
  list.innerHTML = '';
  options.forEach(opt => {
    const option = document.createElement('option');
    option.value = opt;
    list.appendChild(option);
  });
}

