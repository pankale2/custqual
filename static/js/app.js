// Custom Quals Processor - Main JavaScript file

document.addEventListener('DOMContentLoaded', function() {
  // Elements
  const uploadArea = document.getElementById('uploadArea');
  const fileInput = document.getElementById('fileInput');
  const preview = document.getElementById('preview');
  const fileNameEl = document.getElementById('fileName');
  const fileSizeEl = document.getElementById('fileSize');
  const removeFileBtn = document.getElementById('removeFile');
  const processBtn = document.getElementById('process-btn');
  const darkModeToggle = document.getElementById('dark-mode-toggle');
  const processingOverlay = document.getElementById('processing-overlay');

  let isProcessing = false;
  let dragCounter = 0;

  const fmtSize = bytes => { if(!bytes) return '0 Bytes'; const k=1024,s=['Bytes','KB','MB','GB']; const i=Math.floor(Math.log(bytes)/Math.log(k)); return (bytes/Math.pow(k,i)).toFixed(2)+' '+s[i]; };

  // Drag/drop visuals with counter to avoid flicker
  const prevent = e => { e.preventDefault(); e.stopPropagation(); };
  ['dragenter','dragover','dragleave','drop'].forEach(evt => uploadArea.addEventListener(evt, prevent, false));

  uploadArea.addEventListener('dragenter', () => { dragCounter++; uploadArea.classList.add('dragover'); });
  uploadArea.addEventListener('dragover', () => { uploadArea.classList.add('dragover'); });
  uploadArea.addEventListener('dragleave', () => { dragCounter = Math.max(0, dragCounter-1); if(dragCounter===0) uploadArea.classList.remove('dragover'); });
  uploadArea.addEventListener('drop', e => {
    dragCounter = 0;
    uploadArea.classList.remove('dragover');
    const f = e.dataTransfer.files;
    if(f && f.length){ fileInput.files = f; handleFile(); }
  });

  // Keyboard accessibility for upload area
  uploadArea.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
  });

  fileInput.addEventListener('change', handleFile);
  removeFileBtn.addEventListener('click', ()=> { setPreview(null); uploadArea.focus(); });

  function handleFile(){
    const file = fileInput.files[0];
    if(!file) return setPreview(null);
    if(!file.name.toLowerCase().endsWith('.xlsx')){ alert('Please select an Excel file (.xlsx)'); fileInput.value=''; return setPreview(null); }
    setPreview(file);
  }

  function setPreview(file){
    if(!file){
      preview.classList.remove('show');
      preview.style.display='none';
      preview.setAttribute('aria-hidden','true');
      processBtn.disabled=true;
      processBtn.value = 'Process & Download';
      return;
    }
    fileNameEl.textContent = file.name;
    fileNameEl.title = file.name;
    fileSizeEl.textContent = fmtSize(file.size);
    preview.classList.add('show');
    preview.style.display='flex';
    preview.setAttribute('aria-hidden','false');
    processBtn.disabled = false;
    processBtn.value = 'Process & Download';
  }

  function showProcessingOverlay() {
    if (processingOverlay) {
      processingOverlay.style.display = 'flex';
    }
  }
  function hideProcessingOverlay() {
    if (processingOverlay) {
      processingOverlay.style.display = 'none';
    }
  }

  // Process file via fetch -> blob download (guard against double submit)
  document.getElementById('main-form').addEventListener('submit', function(e) {
    e.preventDefault();
    if (isProcessing) return;
    const f = fileInput.files[0];
    if(!f) return alert('Choose a file first');

    isProcessing = true;
    processBtn.value = 'Processing…';
    processBtn.disabled = true;
    showProcessingOverlay();

    const fd = new FormData(); fd.append('file', f, f.name);
    fetch('/', { method: 'POST', body: fd })
    .then(resp => {
      if(!resp.ok) throw new Error('Server error');
      const cd = resp.headers.get('Content-Disposition') || '';
      return resp.blob().then(blob => {
        let filename = f.name;
        const filenameMatch = /filename\*=UTF-8''([^;]+)|filename="([^"]+)"|filename=([^;]+)/.exec(cd);
        if (filenameMatch) {
          const extractedName = filenameMatch[1] || filenameMatch[2] || filenameMatch[3];
          if (extractedName) {
            filename = decodeURIComponent(extractedName.trim());
          }
        }
        if (!filename.includes('_')) {
          const timestamp = new Date().toISOString().slice(0,19).replace(/[-:]/g,'').replace('T','_');
          const baseName = filename.replace(/\.xlsx?$/i, '');
          filename = `${baseName}_${timestamp}.xlsx`;
        }
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a);
        a.click(); a.remove(); URL.revokeObjectURL(url);
        reset();
      });
    })
    .catch(err => { alert('Processing failed: ' + err.message); reset(); });
  });

  function reset(){
    isProcessing = false;
    processBtn.value = 'Process & Download';
    processBtn.disabled = false;
    fileInput.value = '';
    setPreview(null);
    hideProcessingOverlay();
  }

  // Dark mode toggle logic
  function setDarkMode(enabled) {
    document.body.classList.toggle('dark-mode', enabled);
    darkModeToggle.textContent = enabled ? '☀️ Light Mode' : '🌙 Dark Mode';
  }
  darkModeToggle.addEventListener('click', function(e) {
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();
    const isCurrentlyDark = document.body.classList.contains('dark-mode');
    setDarkMode(!isCurrentlyDark);
    localStorage.setItem('dark_mode', !isCurrentlyDark ? '1' : '0');
  });
  // Restore dark mode preference
  const darkPref = localStorage.getItem('dark_mode');
  setDarkMode(darkPref === '1');

  // Show shutdown button only in EXE mode
  function checkEXEMode() {
    const isEXE = window.navigator.userAgent.includes('Electron') || 
                 window.location.protocol === 'file:' ||
                 window.location.hostname === '127.0.0.1';
    if (isEXE) {
      const shutdownBtn = document.getElementById('shutdown-btn');
      if (shutdownBtn) shutdownBtn.style.display = 'block';
    }
  }
  checkEXEMode();

  // Shutdown function
  window.shutdownApp = function() {
    if (confirm('Are you sure you want to exit the application?')) {
      fetch('/shutdown', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
      }).then(() => {
        alert('Application is shutting down. You can close this browser window.');
        window.close();
      }).catch(() => {
        alert('Application is shutting down. You can close this browser window.');
        window.close();
      });
    }
  };
});