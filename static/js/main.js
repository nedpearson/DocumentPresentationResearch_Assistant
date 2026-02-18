/* DocIQ — main.js */

// Safe markdown renderer — always sanitize with DOMPurify before inserting into DOM
function safeMarkdown(mdText) {
  const raw = marked.parse(mdText || '');
  return typeof DOMPurify !== 'undefined' ? DOMPurify.sanitize(raw) : raw;
}

// Document text preview modal
async function openPreview(docId) {
  const modal = document.getElementById('previewModal');
  const body  = document.getElementById('previewModalBody');
  const meta  = document.getElementById('previewModalMeta');
  const title = document.getElementById('previewModalTitle');

  modal.style.display = 'flex';
  body.innerHTML = '<div class="preview-text" style="color:var(--text-muted)">Loading extracted text...</div>';

  try {
    const resp = await fetch(`/api/documents/${docId}/text`);
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || 'Failed to load');

    title.textContent = data.name;
    meta.textContent  = `${data.word_count.toLocaleString()} words · ${data.char_count.toLocaleString()} characters`;

    if (!data.has_text || !data.text) {
      body.innerHTML = `<div class="no-text-warning">
        <svg viewBox="0 0 24 24" width="40" height="40" opacity=".3"><circle cx="12" cy="12" r="10"/>
          <line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
        </svg>
        <p>No extractable text found in this document.<br>It may be an image-based PDF or empty file.</p>
      </div>`;
    } else {
      const escaped = data.text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      body.innerHTML = `<div class="preview-text">${escaped}</div>`;
    }
  } catch (e) {
    body.innerHTML = `<div class="no-text-warning"><p>Error: ${e.message}</p></div>`;
  }
}

function closePreviewModal() {
  document.getElementById('previewModal').style.display = 'none';
}

function closePreview(e) {
  if (e.target === document.getElementById('previewModal')) closePreviewModal();
}

// Close modal with Escape key
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') closePreviewModal();
});

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
