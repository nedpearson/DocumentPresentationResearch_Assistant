/* DocIQ — main.js */

// Theme toggle
function toggleTheme() {
  const html = document.documentElement;
  html.dataset.theme = html.dataset.theme === 'dark' ? 'light' : 'dark';
  localStorage.setItem('theme', html.dataset.theme);
}
(function () {
  const saved = localStorage.getItem('theme');
  if (saved) document.documentElement.dataset.theme = saved;
})();

// Mobile sidebar
function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebarOverlay').classList.toggle('show');
}

// Toast notifications
function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  if (!container) return;
  const icons = { success: '✓', error: '✗', warning: '⚠' };
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `<span>${icons[type] || '•'}</span><span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'slideOut .2s ease forwards';
    setTimeout(() => toast.remove(), 200);
  }, 3500);
}

// File upload helper
async function uploadFiles(files, onComplete) {
  if (!files.length) return;
  let done = 0;
  const results = [];

  for (const file of files) {
    const form = new FormData();
    form.append('file', file);
    try {
      const resp = await fetch('/api/upload', { method: 'POST', body: form });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.error || 'Upload failed');
      results.push(data);
      showToast(`Uploaded: ${data.name}`, 'success');
    } catch (e) {
      showToast(`${file.name}: ${e.message}`, 'error');
    }
    done++;
    if (done === files.length && typeof onComplete === 'function') onComplete(results);
  }
}

// Drag-and-drop zone initializer
function initDropZone(zoneId, inputId, onComplete) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  if (!zone || !input) return;

  zone.addEventListener('click', () => input.click());
  input.addEventListener('change', () => uploadFiles(Array.from(input.files), onComplete));

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = Array.from(e.dataTransfer.files);
    uploadFiles(files, onComplete);
  });
}
