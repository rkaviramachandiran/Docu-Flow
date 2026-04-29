document.addEventListener('DOMContentLoaded', () => {
    // Views
    const uploadView = document.getElementById('upload-view');
    const editorView = document.getElementById('editor-view');
    const successView = document.getElementById('success-view');

    // Upload Elements
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const selectFilesBtn = document.getElementById('select-files-btn');
    const addMoreFilesBtn = document.getElementById('add-more-files-btn');

    // Editor Elements
    const filePreviewArea = document.getElementById('file-preview-area');
    const convertBtn = document.getElementById('convert-btn');
    const convertingOverlay = document.getElementById('converting-overlay');
    const errorOverlay = document.getElementById('error-overlay');
    const errorText = document.getElementById('error-text');
    const retryBtn = document.getElementById('retry-btn');

    // Success Elements
    const startOverBtn = document.getElementById('start-over-btn');

    let currentFiles = [];
    let convertedFiles = [];
    let selectedOutputFormat = document.body.getAttribute('data-mode') || 'pdf';

    // --- Drag and Drop Logic ---
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        handleFiles(e.dataTransfer.files);
    });

    // --- Click Upload Logic ---
    selectFilesBtn.addEventListener('click', () => {
        fileInput.value = '';
        fileInput.click();
    });
    
    if (addMoreFilesBtn) {
        addMoreFilesBtn.addEventListener('click', () => {
            fileInput.value = '';
            fileInput.click();
        });
    }
    
    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        if (files.length === 0) return;
        
        const validExtensions = ['.doc', '.docx', '.xls', '.xlsx', '.png', '.jpg', '.jpeg', '.txt', '.pdf'];
        
        const newFiles = Array.from(files).filter(file => {
            const fileNameStr = file.name.toLowerCase();
            return validExtensions.some(ext => fileNameStr.endsWith(ext));
        });
        
        if (newFiles.length === 0) {
            alert("Invalid file type(s). Please upload Word, Excel, Images, Text or PDF documents.");
            return;
        }

        // Add to array
        currentFiles = [...currentFiles, ...newFiles];
        showEditorView();
    }

    // --- View Transitions ---
    function showEditorView() {
        filePreviewArea.innerHTML = ''; // Clear
        
        currentFiles.forEach((file, index) => {
            const ext = file.name.split('.').pop().toLowerCase();
            const isWord = ext.startsWith('doc');
            const isExcel = ext.startsWith('xls');
            const isImage = ['png', 'jpg', 'jpeg'].includes(ext);
            const isText = ext === 'txt';
            const isPdf = ext === 'pdf';
            
            let iconClass = 'fa-file-pdf';
            let iconColor = '#ef4444'; // Red for PDF
            
            if (isWord) {
                iconClass = 'fa-file-word';
                iconColor = '#4f46e5';
            } else if (isExcel) {
                iconClass = 'fa-file-excel';
                iconColor = '#10b981';
            } else if (isImage) {
                iconClass = 'fa-file-image';
                iconColor = '#f59e0b';
            } else if (isText) {
                iconClass = 'fa-file-lines';
                iconColor = '#64748b';
            }
            
            const card = document.createElement('div');
            card.className = 'file-card';
            
            card.innerHTML = `
                <div class="file-card-preview">
                    <i class="fa-solid ${iconClass}" style="color: ${iconColor}; font-size: 80px;"></i>
                </div>
                <div class="file-card-name" title="${file.name}">${file.name}</div>
                <button class="remove-file-btn tooltip" data-index="${index}" data-tooltip="Remove file">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            `;
            filePreviewArea.appendChild(card);
        });

        // Attach listeners to new buttons
        document.querySelectorAll('.remove-file-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.currentTarget.getAttribute('data-index'));
                currentFiles.splice(idx, 1);
                if (currentFiles.length === 0) {
                    resetUpload();
                } else {
                    showEditorView(); // re-render
                }
            });
        });

        // Switch views
        uploadView.classList.add('hidden');
        editorView.classList.remove('hidden');
        successView.classList.add('hidden');
        
        // Reset Overlays
        convertingOverlay.classList.add('hidden');
        errorOverlay.classList.add('hidden');
        convertBtn.disabled = false;
        convertBtn.style.opacity = '1';
    }

    function showSuccessView(files) {
        editorView.classList.add('hidden');
        successView.classList.remove('hidden');
        
        const successTitle = document.querySelector('.success-title');
        const isTargetWord = selectedOutputFormat === 'docx';
        
        if (files.length > 1) {
            successTitle.textContent = isTargetWord ? "Your Word documents are ready!" : "Your PDFs are ready!";
        } else {
            successTitle.textContent = isTargetWord ? "Your Word document is ready!" : "Your PDF is ready!";
        }

        const downloadList = document.getElementById('download-list');
        downloadList.innerHTML = ''; // clear

        files.forEach(file => {
            const btn = document.createElement('button');
            btn.className = 'btn-massive btn-gradient';
            btn.style.width = '100%';
            btn.style.justifyContent = 'space-between';
            btn.innerHTML = `
                <div style="display: flex; align-items: center; gap: 10px; overflow: hidden;">
                    <i class="fa-solid fa-file-pdf"></i>
                    <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 300px;">${file.name}</span>
                </div>
                <i class="fa-solid fa-download"></i>
            `;
            btn.addEventListener('click', () => {
                window.location.href = `/download/${file.file_id}?name=${encodeURIComponent(file.name)}`;
            });
            downloadList.appendChild(btn);
        });
    }

    function resetUpload() {
        currentFiles = [];
        convertedFiles = [];
        fileInput.value = '';
        
        uploadView.classList.remove('hidden');
        editorView.classList.add('hidden');
        successView.classList.add('hidden');
    }

    // --- Actions ---
    startOverBtn.addEventListener('click', resetUpload);

    convertBtn.addEventListener('click', async () => {
        if (currentFiles.length === 0) return;

        convertingOverlay.classList.remove('hidden');
        const overlayText = convertingOverlay.querySelector('p');
        if (overlayText) {
            overlayText.textContent = selectedOutputFormat === 'docx' ? 'Generating Word Document...' : 'Generating PDF...';
        }
        
        errorOverlay.classList.add('hidden');
        convertBtn.disabled = true;
        convertBtn.style.opacity = '0.5';

        const formData = new FormData();
        formData.append('output_format', selectedOutputFormat);
        // Append multiple files
        currentFiles.forEach(file => {
            formData.append('files', file);
        });

        try {
            const response = await fetch('/convert', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.detail || 'Conversion failed');
            }

            convertedFiles = data.files;
            showSuccessView(convertedFiles);
            
        } catch (error) {
            convertingOverlay.classList.add('hidden');
            errorOverlay.classList.remove('hidden');
            errorText.textContent = error.message;
            convertBtn.disabled = false;
            convertBtn.style.opacity = '1';
        }
    });

    retryBtn.addEventListener('click', () => {
        errorOverlay.classList.add('hidden');
        convertBtn.click();
    });
});
