// AetherPDF App JavaScript Logic - Python Backend-Powered Conversions

// Configure Tailwind CSS prior to execution
window.tailwind = {
  config: {
    darkMode: "class",
    theme: {
      extend: {
        "colors": {
          "on-tertiary-fixed": "#001e2f",
          "on-tertiary": "#00344d",
          "on-background": "#e5e2e1",
          "on-tertiary-fixed-variant": "#004c6e",
          "error-container": "#93000a",
          "on-tertiary-container": "#e4f2ff",
          "on-secondary-fixed": "#2c0051",
          "surface-container": "#201f1f",
          "on-secondary-fixed-variant": "#6900b3",
          "surface-dim": "#131313",
          "surface-container-high": "#2a2a2a",
          "surface-tint": "#b8c3ff",
          "surface-container-lowest": "#0e0e0e",
          "background": "#131313",
          "secondary-container": "#6f00be",
          "surface-container-low": "#1c1b1b",
          "on-primary-container": "#efefff",
          "surface-variant": "#353534",
          "tertiary": "#89ceff",
          "on-error": "#690005",
          "inverse-primary": "#124af0",
          "surface-bright": "#3a3939",
          "secondary-fixed": "#f0dbff",
          "inverse-on-surface": "#313030",
          "secondary-fixed-dim": "#ddb7ff",
          "primary-fixed-dim": "#b8c3ff",
          "on-secondary-container": "#d6a9ff",
          "surface": "#131313",
          "secondary": "#ddb7ff",
          "on-primary": "#002388",
          "inverse-surface": "#e5e2e1",
          "tertiary-fixed-dim": "#89ceff",
          "tertiary-container": "#0074a6",
          "tertiary-fixed": "#c9e6ff",
          "on-error-container": "#ffdad6",
          "error": "#ffb4ab",
          "primary-container": "#2e5bff",
          "on-surface": "#e5e2e1",
          "primary-fixed": "#dde1ff",
          "on-primary-fixed": "#001356",
          "outline-variant": "#434656",
          "on-surface-variant": "#c4c5d9",
          "on-primary-fixed-variant": "#0035be",
          "primary": "#b8c3ff",
          "surface-container-highest": "#353534",
          "on-secondary": "#490080",
          "outline": "#8e90a2"
        },
        "borderRadius": {
          "DEFAULT": "0.25rem",
          "lg": "0.5rem",
          "xl": "0.75rem",
          "full": "9999px"
        },
        "spacing": {
          "unit": "4px",
          "stack-md": "1rem",
          "stack-lg": "2rem",
          "container-padding": "2rem",
          "gutter": "1.5rem",
          "stack-sm": "0.5rem"
        },
        "fontFamily": {
          "display-lg-mobile": ["Geist"],
          "body-md": ["Geist"],
          "body-lg": ["Geist"],
          "headline-md": ["Geist"],
          "display-lg": ["Geist"],
          "label-caps": ["JetBrains Mono"],
          "code-sm": ["JetBrains Mono"]
        },
        "fontSize": {
          "display-lg-mobile": ["32px", { "lineHeight": "1.2", "letterSpacing": "-0.02em", "fontWeight": "700" }],
          "body-md": ["16px", { "lineHeight": "1.5", "fontWeight": "400" }],
          "body-lg": ["18px", { "lineHeight": "1.6", "fontWeight": "400" }],
          "headline-md": ["24px", { "lineHeight": "1.3", "letterSpacing": "-0.01em", "fontWeight": "600" }],
          "display-lg": ["48px", { "lineHeight": "1.1", "letterSpacing": "-0.04em", "fontWeight": "700" }],
          "label-caps": ["12px", { "lineHeight": "1.0", "letterSpacing": "0.1em", "fontWeight": "500" }],
          "code-sm": ["13px", { "lineHeight": "1.4", "fontWeight": "400" }]
        }
      }
    }
  }
};

document.addEventListener('DOMContentLoaded', () => {
  // --- UI Elements ---
  const navConverterBtn = document.getElementById('nav-converter-btn');
  const navToolsBtn = document.getElementById('nav-tools-btn');
  const logoBtn = document.getElementById('logo-btn');
  const subnavWorkspace = document.getElementById('subnav-workspace');
  
  const dashboardView = document.getElementById('dashboard-view');
  const converterWorkspace = document.getElementById('converter-workspace');
  const toolsSection = document.getElementById('tools-section');
  const backToDashboardBtn = document.getElementById('back-to-dashboard-btn');
  const activeConverterBadge = document.getElementById('active-converter-badge');
  const allowedFormatsInfo = document.getElementById('allowed-formats-info');
  
  const uploadZone = document.getElementById('upload-zone');
  const fileInput = document.getElementById('file-input');
  
  const fileSelectedState = document.getElementById('file-selected-state');
  const selectedFilename = document.getElementById('selected-filename');
  const selectedFilesize = document.getElementById('selected-filesize');
  const selectedFileIcon = document.getElementById('selected-file-icon');
  const removeFileBtn = document.getElementById('remove-file-btn');
  const startConversionBtn = document.getElementById('start-conversion-btn');
  const actionBtnText = document.getElementById('action-btn-text');
  
  const processingState = document.getElementById('processing-state');
  const progressCircle = document.getElementById('progress-circle');
  const progressPercentageText = document.getElementById('progress-percentage-text');
  const consoleLogs = document.getElementById('console-logs');
  
  const successState = document.getElementById('success-state');
  const outputFilenameText = document.getElementById('output-filename');
  const successFileIcon = document.getElementById('success-file-icon');
  const downloadOutputBtn = document.getElementById('download-output-btn');
  const convertAnotherBtn = document.getElementById('convert-another-btn');
  const recentTransmutationsList = document.getElementById('recent-transmutations-list');

  // --- State Variables ---
  let activeMode = null; // 'pdf-to-doc', 'pdf-to-ppt', etc.
  let selectedFile = null;
  let convertedBlob = null;
  let convertedFilename = '';
  const recentTransmutations = {}; // stores successfully converted blobs

  const modeSettings = {
    'pdf-to-doc': {
      title: 'PDF to DOCX',
      accept: '.pdf',
      mime: 'application/pdf',
      actionText: 'TRANSMUTE TO DOCX',
      icon: 'picture_as_pdf',
      outIcon: 'description',
      iconColorClass: 'text-primary bg-primary/10 border-primary/20',
      ext: 'docx'
    },
    'pdf-to-ppt': {
      title: 'PDF to PPTX',
      accept: '.pdf',
      mime: 'application/pdf',
      actionText: 'TRANSMUTE TO PPTX',
      icon: 'picture_as_pdf',
      outIcon: 'slideshow',
      iconColorClass: 'text-orange-400 bg-orange-500/10 border-orange-500/20',
      ext: 'pptx'
    },
    'pdf-to-sheets': {
      title: 'PDF to XLSX',
      accept: '.pdf',
      mime: 'application/pdf',
      actionText: 'TRANSMUTE TO XLSX',
      icon: 'picture_as_pdf',
      outIcon: 'table_chart',
      iconColorClass: 'text-green-400 bg-green-500/10 border-green-500/20',
      ext: 'xlsx'
    },
    'doc-to-pdf': {
      title: 'DOCX to PDF',
      accept: '.docx',
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      actionText: 'TRANSMUTE TO PDF',
      icon: 'description',
      outIcon: 'picture_as_pdf',
      iconColorClass: 'text-secondary bg-secondary/10 border-secondary/20',
      ext: 'pdf'
    },
    'ppt-to-pdf': {
      title: 'PPTX to PDF',
      accept: '.pptx',
      mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      actionText: 'TRANSMUTE TO PDF',
      icon: 'slideshow',
      outIcon: 'picture_as_pdf',
      iconColorClass: 'text-secondary bg-secondary/10 border-secondary/20',
      ext: 'pdf'
    },
    'sheets-to-pdf': {
      title: 'XLSX to PDF',
      accept: '.xlsx',
      mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      actionText: 'TRANSMUTE TO PDF',
      icon: 'table_chart',
      outIcon: 'picture_as_pdf',
      iconColorClass: 'text-secondary bg-secondary/10 border-secondary/20',
      ext: 'pdf'
    }
  };

  // --- Navigation & View Switching ---
  function showConverterView() {
    navConverterBtn.className = "font-body-md text-body-md text-primary border-b-2 border-primary pb-1";
    navToolsBtn.className = "font-body-md text-body-md text-on-surface-variant hover:text-on-surface transition-colors";
    toolsSection.classList.add('hidden');
    if (activeMode) {
      dashboardView.classList.add('hidden');
      converterWorkspace.classList.remove('hidden');
    } else {
      dashboardView.classList.remove('hidden');
      converterWorkspace.classList.add('hidden');
    }
  }

  function showToolsView() {
    navConverterBtn.className = "font-body-md text-body-md text-on-surface-variant hover:text-on-surface transition-colors";
    navToolsBtn.className = "font-body-md text-body-md text-primary border-b-2 border-primary pb-1";
    dashboardView.classList.add('hidden');
    converterWorkspace.classList.add('hidden');
    toolsSection.classList.remove('hidden');
  }

  navConverterBtn.addEventListener('click', (e) => {
    e.preventDefault();
    showConverterView();
  });
  
  navToolsBtn.addEventListener('click', (e) => {
    e.preventDefault();
    showToolsView();
  });

  logoBtn.addEventListener('click', () => {
    activeMode = null;
    showConverterView();
  });

  subnavWorkspace.addEventListener('click', (e) => {
    e.preventDefault();
    activeMode = null;
    showConverterView();
  });

  // Grid cards interaction
  document.querySelectorAll('[data-mode]').forEach(card => {
    card.addEventListener('click', () => {
      activeMode = card.getAttribute('data-mode');
      const settings = modeSettings[activeMode];
      
      activeConverterBadge.textContent = settings.title;
      allowedFormatsInfo.textContent = `ACCEPTED FORMATS: ${settings.accept.toUpperCase()}`;
      actionBtnText.textContent = settings.actionText;
      
      const wrapper = document.getElementById('file-icon-wrapper');
      wrapper.className = `w-14 h-14 rounded-xl flex items-center justify-center ${settings.iconColorClass}`;
      selectedFileIcon.textContent = settings.icon;
      
      resetWorkspaceStates();
      
      dashboardView.classList.add('hidden');
      converterWorkspace.classList.remove('hidden');
    });
  });

  backToDashboardBtn.addEventListener('click', () => {
    activeMode = null;
    showConverterView();
  });

  // --- Upload Handlers ---
  uploadZone.addEventListener('click', () => {
    fileInput.accept = modeSettings[activeMode].accept;
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  // Drag-and-drop animations
  ['dragenter', 'dragover'].forEach(eventName => {
    uploadZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      uploadZone.classList.add('border-primary/50', 'bg-primary/5');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    uploadZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      uploadZone.classList.remove('border-primary/50', 'bg-primary/5');
    }, false);
  });

  uploadZone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleFileSelected(files[0]);
    }
  });

  function handleFileSelected(file) {
    const settings = modeSettings[activeMode];
    const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
    
    if (ext !== settings.accept) {
      alert(`Invalid format! Please upload a ${settings.accept} file.`);
      return;
    }

    if (file.size > 50 * 1024 * 1024) {
      alert('File exceeds the 50MB sandbox limit.');
      return;
    }

    selectedFile = file;
    selectedFilename.textContent = file.name;
    selectedFilesize.textContent = formatBytes(file.size);

    uploadZone.classList.add('hidden');
    fileSelectedState.classList.remove('hidden');
  }

  removeFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    selectedFile = null;
    fileInput.value = '';
    fileSelectedState.classList.add('hidden');
    uploadZone.classList.remove('hidden');
  });

  // --- Terminal Logging ---
  function addLog(text, type = 'info') {
    const timeStr = new Date().toLocaleTimeString();
    const logLine = document.createElement('div');
    
    const timeSpan = document.createElement('span');
    timeSpan.style.color = '#64748b'; // slate grey
    timeSpan.style.marginRight = '8px';
    timeSpan.textContent = `[${timeStr}]`;
    
    const textSpan = document.createElement('span');
    if (type === 'success') {
      textSpan.style.color = '#4ade80'; // emerald green
    } else if (type === 'warn') {
      textSpan.style.color = '#f97316'; // orange
    } else if (type === 'error') {
      textSpan.style.color = '#f87171'; // red
    } else {
      textSpan.style.color = '#89ceff'; // tertiary blue
    }
    textSpan.textContent = text;
    
    logLine.appendChild(timeSpan);
    logLine.appendChild(textSpan);
    consoleLogs.appendChild(logLine);
    consoleLogs.scrollTop = consoleLogs.scrollHeight;
  }

  function clearLogs() {
    consoleLogs.innerHTML = '';
  }

  function setProgress(percent) {
    const radius = 50;
    const circumference = 2 * Math.PI * radius;
    const offset = circumference - (percent / 100) * circumference;
    progressCircle.style.strokeDashoffset = offset;
    progressPercentageText.textContent = `${Math.round(percent)}%`;
  }

  function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  function resetWorkspaceStates() {
    uploadZone.classList.remove('hidden');
    fileSelectedState.classList.add('hidden');
    processingState.classList.add('hidden');
    successState.classList.add('hidden');
    fileInput.value = '';
    clearLogs();
    setProgress(0);
    convertedBlob = null;
  }

  // --- Python Backend Conversion Hook ---
  startConversionBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    fileSelectedState.classList.add('hidden');
    processingState.classList.remove('hidden');
    clearLogs();
    setProgress(0);

    addLog(`Initializing Python Neural Engine...`);
    addLog(`Payload: ${selectedFile.name} (${formatBytes(selectedFile.size)})`);
    addLog(`Connecting to local conversion server at http://localhost:5000/convert...`);

    // Simulate progress increments during backend processing
    let progress = 0;
    const progressInterval = setInterval(() => {
      if (progress < 90) {
        progress += Math.floor(Math.random() * 5) + 2;
        if (progress > 90) progress = 90;
        setProgress(progress);
        
        if (progress > 20 && progress < 40 && consoleLogs.children.length === 3) {
          addLog(`[Python] Worker active. Reading binary file stream...`, 'info');
        } else if (progress > 45 && progress < 60 && consoleLogs.children.length === 4) {
          addLog(`[Python] Running conversion engine for ${activeMode.toUpperCase()}...`, 'info');
          if (activeMode.endsWith('pdf')) {
            addLog(`[Python] Dispatching Microsoft Office COM client...`, 'info');
          } else {
            addLog(`[Python] Reconstructing layout segments and tables...`, 'info');
          }
        } else if (progress > 70 && progress < 85 && consoleLogs.children.length === 6) {
          addLog(`[Python] Rendering binary payload...`, 'info');
        }
      }
    }, 200);

    try {
      const formData = new FormData();
      formData.append('file', selectedFile);
      formData.append('mode', activeMode);

      const response = await fetch('http://localhost:5000/convert', {
        method: 'POST',
        body: formData
      });

      clearInterval(progressInterval);

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error || `Server responded with status ${response.status}`);
      }

      const outputBlob = await response.blob();
      setProgress(100);
      addLog(`[Python] Conversion successful! Transmuted file size: ${formatBytes(outputBlob.size)}`, 'success');

      convertedBlob = outputBlob;
      const baseName = selectedFile.name.substring(0, selectedFile.name.lastIndexOf('.'));
      const ext = modeSettings[activeMode].ext;
      convertedFilename = `${baseName}_converted.${ext}`;

      setTimeout(() => {
        processingState.classList.add('hidden');
        successState.classList.remove('hidden');
        
        outputFilenameText.textContent = convertedFilename;
        successFileIcon.textContent = modeSettings[activeMode].outIcon;

        // Add to recent transmutations history
        addTransmutationHistory(selectedFile.name, convertedFilename, activeMode, convertedBlob);
        
        // Trigger auto download
        downloadFile();
      }, 800);

    } catch (err) {
      clearInterval(progressInterval);
      console.error(err);
      addLog(`[Python Error] ${err.message}`, 'error');
      addLog(`Please ensure python app.py is running on port 5000 and MS Office is installed.`, 'warn');
      setTimeout(() => {
        alert(`Failed: ${err.message}`);
        resetWorkspaceStates();
      }, 5000);
    }
  });

  function downloadFile() {
    if (!convertedBlob) return;
    const url = URL.createObjectURL(convertedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = convertedFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    addLog(`Download trigger complete!`, 'success');
  }

  downloadOutputBtn.addEventListener('click', downloadFile);
  convertAnotherBtn.addEventListener('click', resetWorkspaceStates);

  // --- Dynamic Recent Transmutations Manager ---
  function addTransmutationHistory(srcName, outName, mode, blob) {
    const id = 'transmute_' + Date.now();
    recentTransmutations[id] = {
      blob: blob,
      filename: outName
    };

    const ext = modeSettings[mode].ext.toUpperCase();
    const row = document.createElement('div');
    row.className = 'glass-card p-4 rounded-xl flex items-center justify-between border-white/5 opacity-0 translate-y-2 transition-all duration-500';
    
    let iconClass = 'text-primary';
    if (ext === 'PPTX') iconClass = 'text-orange-400';
    if (ext === 'XLSX') iconClass = 'text-green-400';
    if (ext === 'PDF') iconClass = 'text-secondary';

    row.innerHTML = `
      <div class="flex items-center gap-4">
        <div class="w-10 h-10 rounded-lg bg-surface-container-high flex items-center justify-center">
          <span class="material-symbols-outlined ${iconClass}">${modeSettings[mode].outIcon}</span>
        </div>
        <div class="text-left">
          <h4 class="font-body-md text-on-surface font-medium max-w-[250px] sm:max-w-[400px] truncate">${outName}</h4>
          <p class="text-[10px] text-on-surface-variant font-label-caps uppercase">Converted from ${ext} • Just now</p>
        </div>
      </div>
      <div class="flex items-center gap-3">
        <span class="text-xs text-green-400 font-label-caps">SUCCESS</span>
        <button class="p-2 hover:bg-white/10 rounded-lg transition-colors material-symbols-outlined text-[20px] download-recent-btn" data-id="${id}">download</button>
      </div>
    `;

    recentTransmutationsList.insertBefore(row, recentTransmutationsList.firstChild);
    
    setTimeout(() => {
      row.classList.remove('opacity-0', 'translate-y-2');
    }, 50);

    row.querySelector('.download-recent-btn').addEventListener('click', (e) => {
      const fileId = e.target.getAttribute('data-id');
      const fileData = recentTransmutations[fileId];
      if (fileData) {
        const url = URL.createObjectURL(fileData.blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileData.filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
    });
  }
});
