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
  const uploadIcon = document.getElementById('uploadIcon');

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
  removeFileBtn.addEventListener('click', (e) => {
    e.preventDefault(); // Prevent default behavior
    fileInput.value = ''; // Clear the file input
    setPreview(null); // Reset the preview area
  });

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

  // Dark mode helpers + upload icon update
  function updateUploadIcon() {
    if (!uploadIcon) return console.error('uploadIcon missing; cannot update icon');
    // Always use uploadw.png for both light and dark modes
    uploadIcon.src = '/static/uploadw.png';
  }

  function setDarkMode(enabled) {
    document.body.classList.toggle('dark-mode', enabled);
    darkModeToggle.textContent = enabled ? '☀️ Light Mode' : '🌙 Dark Mode';
    // persist preference here (single source of truth)
    localStorage.setItem('dark_mode', enabled ? '1' : '0');
    updateUploadIcon();
  }

  // Restore dark mode preference and initialize icon
  const darkPref = localStorage.getItem('dark_mode');
  setDarkMode(darkPref === '1');

  // Toggle via button (use setDarkMode so icon updates)
  darkModeToggle.addEventListener('click', function(e) {
    e.preventDefault();
    e.stopPropagation();
    const currentlyDark = document.body.classList.contains('dark-mode');
    setDarkMode(!currentlyDark);
  });

  // Show shutdown button
  const shutdownBtn = document.getElementById('shutdown-btn');
  if (shutdownBtn) shutdownBtn.style.display = 'block';

  // Shutdown function
  window.shutdownApp = function() {
    fetch('http://localhost:5001/shutdown', { method: 'POST' }) // Updated port to 5001
        .then(response => {
            if (response.ok) {
                console.log('Application is shutting down.');
                // Small delay to ensure shutdown request is processed
                setTimeout(() => {
                    window.close(); // Attempt to close the browser tab
                    // Fallback: redirect to a blank page if window.close() doesn't work
                    setTimeout(() => {
                        window.location.href = 'about:blank';
                    }, 500);
                }, 100);
            } else {
                console.error('Failed to shut down the application.');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            // Even if there's an error, try to close the tab
            window.close();
        });
  };

  // Debugging logic
  console.log('Debugging app.js');

  // Check if uploadIcon exists
  if (!uploadIcon) {
      console.error('uploadIcon element not found');
  }

  // Check if dark-mode-toggle exists
  if (!darkModeToggle) {
      console.error('dark-mode-toggle element not found');
  }

  // Debugging updateUploadIcon function
  function debugUpdateUploadIcon() {
      console.log('Using uploadw.png for all modes');
      uploadIcon.src = '/static/uploadw.png';
  }

  debugUpdateUploadIcon();
});