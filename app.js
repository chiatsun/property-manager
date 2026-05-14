const API_URL = "https://script.google.com/macros/s/AKfycbyJkPfvNA0gosQeGaN7INsRVe82P-fCkN4ZWSenHSviMh-6pUYv8IB4vEmStYiYwPLg1w/exec";

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
  document.getElementById('detail-modal').classList.add('active');
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
