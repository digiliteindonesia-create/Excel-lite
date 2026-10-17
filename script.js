let currentFilter = 'all';
let completedFormulas = JSON.parse(localStorage.getItem('excel_completed')) || [];

window.onload = function() {
  renderCards(formulasData);
  if (localStorage.getItem('excel_logged_in') === 'true') {
    applyLoggedInState();
  } else {
    applyLoggedOutState();
  }
};

function applyLoggedInState() {
  document.getElementById('mainContent').classList.remove('is-locked');
  document.getElementById('lockBanner').style.display = 'none';
  document.getElementById('telegramBtn').classList.remove('is-locked');
  document.getElementById('loginRightPanel').classList.add('hidden');
  document.getElementById('mainLayout').classList.add('is-logged-in');
  document.getElementById('logoutBtn').style.display = 'flex';
}

function applyLoggedOutState() {
  document.getElementById('mainContent').classList.add('is-locked');
  document.getElementById('lockBanner').style.display = 'flex';
  document.getElementById('telegramBtn').classList.add('is-locked');
  document.getElementById('loginRightPanel').classList.remove('hidden');
  document.getElementById('mainLayout').classList.remove('is-logged-in');
  document.getElementById('logoutBtn').style.display = 'none';
}

function handleLogin(e) {
  e.preventDefault();
  localStorage.setItem('excel_logged_in', 'true');
  applyLoggedInState();
  showToast("Berhasil login!");
}

function handleLogout() {
  localStorage.removeItem('excel_logged_in');
  applyLoggedOutState();
  showToast("Anda telah keluar.");
}

function renderCards(data) {
  const grid = document.getElementById('cardGrid');
  grid.innerHTML = '';

  data.forEach(item => {
    const isDone = completedFormulas.includes(item.id);
    const card = document.createElement('div');
    card.className = `card ${item.level}`;
    
    card.innerHTML = `
      <div class="card-header">
        <h3>${item.name}</h3>
        <span class="badge">${item.levelText}</span>
      </div>
      <p style="font-size: 0.9rem; color: var(--text-muted);">${item.desc}</p>
      <div class="syntax-box"><span>Sintaks</span>${item.syntax}</div>
      
      <!-- Tombol Upload Jawaban -->
      <div class="upload-zone">
        <label class="upload-label" for="upload-${item.id}">
          <i class="fas fa-upload"></i> ${isDone ? 'Jawaban Benar ✅ (Upload Ulang)' : 'Upload & Cek Jawaban'}
        </label>
        <input type="file" id="upload-${item.id}" class="file-input" accept=".xlsx, .xls" onchange="checkAnswer('${item.id}', event)">
      </div>

      <div class="card-footer">
        <button class="btn btn-download" onclick="downloadExcel('${item.name}')"><i class="fas fa-download"></i> Soal</button>
      </div>
    `;
    grid.appendChild(card);
  });

  updateCounters();
  updateProgress();
}

// Logika membaca & mengecek jawaban Excel pakai SheetJS
function checkAnswer(id, event) {
  const file = event.target.files[0];
  if (!file) return;

  const item = formulasData.find(f => f.id === id);
  if (!item) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Mengambil nilai sel jawaban yang ditentukan di formulas.js
      const cell = worksheet[item.expectedCell];
      const userValue = cell ? cell.v : null;

      if (userValue == item.expectedValue) {
        showToast(`🎉 Jawaban ${item.name} BENAR! Selamat!`);
        if (!completedFormulas.includes(id)) {
          completedFormulas.push(id);
          localStorage.setItem('excel_completed', JSON.stringify(completedFormulas));
          renderCards(formulasData);
        }
      } else {
        showToast(`❌ Jawaban ${item.name} masih SALAH. Nilai ditemukan: ${userValue || 'Kosong'}`);
      }
    } catch (err) {
      showToast("Gagal membaca file Excel. Pastikan format file .xlsx valid.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function updateProgress() {
  const total = formulasData.length;
  const done = completedFormulas.length;
  const pct = Math.round((done / total) * 100);
  document.getElementById('progressText').innerText = `${done} dari ${total} Selesai (${pct}%)`;
  document.getElementById('progressBarFill').style.width = `${pct}%`;
}

function updateCounters() {
  document.getElementById('count-all').innerText = formulasData.length;
  document.getElementById('count-basic').innerText = formulasData.filter(f => f.level === 'basic').length;
  document.getElementById('count-intermediate').innerText = formulasData.filter(f => f.level === 'intermediate').length;
  document.getElementById('count-advanced').innerText = formulasData.filter(f => f.level === 'advanced').length;
}

function filterLevel(lvl, e) {
  currentFilter = lvl;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  if (e) e.currentTarget.classList.add('active');
  const filtered = lvl === 'all' ? formulasData : formulasData.filter(f => f.level === lvl);
  renderCards(filtered);
}

function searchFormulas() {
  const kw = document.getElementById('searchInput').value.toLowerCase();
  const filtered = formulasData.filter(f => f.name.toLowerCase().includes(kw) || f.desc.toLowerCase().includes(kw));
  renderCards(filtered);
}

function downloadExcel(name) {
  showToast(`Mengunduh file soal ${name}...`);
}

function showToast(msg) {
  const toast = document.getElementById('toastBox');
  document.getElementById('toastMsg').innerText = msg;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), 3500);
}