// State
let parsedFiles = [];

// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const uploadButton = document.getElementById('uploadButton');
const fileInput = document.getElementById('fileInput');
const resultsSection = document.getElementById('resultsSection');
const resultsBody = document.getElementById('resultsBody');
const resultsTable = document.querySelector('.results-table');
const emptyState = document.getElementById('emptyState');
const loadingIndicator = document.getElementById('loadingIndicator');

// Web Worker code as a string
const workerCode = `
importScripts('https://unpkg.com/jszip@3.10.1/dist/jszip.min.js');
importScripts('https://unpkg.com/sax@1.2.4/lib/sax.js');

console.log('Worker started. JSZip:', typeof JSZip, 'sax:', typeof sax);

self.addEventListener('message', async (event) => {
    const { file, fileName } = event.data;

    try {
        const zip = await JSZip.loadAsync(file);
        
        // First try standard Excel path
        let sheetEntry = zip.file("xl/worksheets/sheet1.xml");
        
        // If not found, check if this is a ZIP containing an XLSX
        if (!sheetEntry) {
            // Look for any .xlsx file inside the zip
            const xlsxFiles = Object.keys(zip.files).filter(path => path.endsWith('.xlsx') && !path.startsWith('__MACOSX'));
            
            if (xlsxFiles.length > 0) {
                console.log('Found nested XLSX: ' + xlsxFiles[0]);
                // Read the inner xlsx file
                const innerXlsxData = await zip.file(xlsxFiles[0]).async("arraybuffer");
                // Load it as a new zip
                const innerZip = await JSZip.loadAsync(innerXlsxData);
                // Now look for sheet1 in the inner zip
                sheetEntry = innerZip.file("xl/worksheets/sheet1.xml");
            } else {
                // Fallback: look for ANY xml file (e.g. for e-Ledger files)
                const xmlFiles = Object.keys(zip.files).filter(path => path.endsWith('.xml') && !path.startsWith('_rels') && !path.includes('/_rels/'));
                if (xmlFiles.length > 0) {
                    sheetEntry = zip.file(xmlFiles[0]);
                    console.log('Standard Excel path not found. Using found XML: ' + xmlFiles[0]);
                }
            }
        }
        
        if (!sheetEntry) {
            const fileList = Object.keys(zip.files).join(", ");
            throw new Error('Excel dosyası veya geçerli bir yapı bulunamadı. İçerik: ' + fileList);
        }

        const xmlText = await sheetEntry.async("string");
        const parser = sax.parser(true); // strict = true
        
        const targets = ["A2", "B3", "B4"];
        const results = { "A2": "", "B3": "", "B4": "" };
        
        let currentCell = null;
        let readingValue = false;

        parser.onopentag = (node) => {
            if (node.name === "c") {
                const r = node.attributes.r;
                if (targets.includes(r)) {
                    currentCell = r;
                }
            }
            // Logic from snippet: insideTargetCell && node.name === "v"
            // Also adding "t" for inlineStr support to be robust
            if (currentCell && (node.name === "v" || node.name === "t")) {
                readingValue = true;
            }
        };

        parser.ontext = (text) => {
            if (readingValue && currentCell) {
                results[currentCell] += text;
            }
        };

        parser.onclosetag = (name) => {
            if (name === "v" || name === "t") {
                readingValue = false;
            }
            if (name === "c") {
                currentCell = null;
            }
        };

        parser.write(xmlText).close();

        self.postMessage({
            success: true,
            data: {
                fileName: fileName,
                title: results["A2"],
                period: results["B3"],
                dateRange: results["B4"]
            }
        });

    } catch (error) {
        self.postMessage({
            success: false,
            error: error.message || 'Dosya işlenirken hata oluştu'
        });
    }
});
`;

// Create a Blob from the worker code once
const workerBlob = new Blob([workerCode], { type: 'application/javascript' });
const workerUrl = URL.createObjectURL(workerBlob);

// Factory function to create a new worker
function createWorker() {
    return new Worker(workerUrl);
}

// File Upload Handlers
uploadButton.addEventListener('click', () => {
    fileInput.click();
});

uploadArea.addEventListener('click', (e) => {
    if (e.target === uploadArea || e.target.closest('.upload-icon, .upload-title, .upload-text')) {
        fileInput.click();
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFiles(Array.from(e.target.files));
    }
    fileInput.value = '';
});

// Drag and Drop Handlers
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');

    if (e.dataTransfer.files.length > 0) {
        handleFiles(Array.from(e.dataTransfer.files));
    }
});

// Handle Multiple Files
function handleFiles(files) {
    // Filter valid files
    const validFiles = files.filter(file => {
        const isValid = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            file.type === 'application/vnd.ms-excel' ||
            file.type === 'application/zip' ||
            file.type === 'application/x-zip-compressed' ||
            file.name.endsWith('.xlsx') ||
            file.name.endsWith('.xls') ||
            file.name.endsWith('.zip');

        if (!isValid) {
            showNotification(`${file.name} geçerli bir Excel dosyası değil.`, 'error');
        }
        return isValid;
    });

    if (validFiles.length === 0) return;

    // Show results section if hidden
    resultsTable.classList.add('active');
    emptyState.classList.add('hidden');

    // Process each file in parallel
    validFiles.forEach(processFile);
}

// Process Single File
function processFile(file) {
    // Create a temporary ID for this file's processing
    const tempId = Date.now() + Math.random().toString(36).substr(2, 9);
    const startTime = Date.now();

    // Add initial row to grid
    const initialData = {
        id: tempId,
        fileName: file.name,
        title: '...',
        period: '...',
        dateRange: '...',
        processingTime: 'İşleniyor...',
        status: 'processing'
    };

    addOrUpdateGridRow(initialData);

    // Create and setup worker
    const worker = createWorker();

    worker.onmessage = (event) => {
        const endTime = Date.now();
        const duration = endTime - startTime;

        if (event.data.success) {
            const resultData = {
                id: tempId,
                ...event.data.data,
                processingTime: duration,
                status: 'success'
            };
            addOrUpdateGridRow(resultData);
            showNotification(`${file.name} tamamlandı (${duration}ms)`, 'success');
        } else {
            const errorData = {
                id: tempId,
                fileName: file.name,
                title: '-',
                period: '-',
                dateRange: '-',
                processingTime: 'Hata',
                status: 'error'
            };
            addOrUpdateGridRow(errorData);
            showNotification(`${file.name} hatası: ${event.data.error}`, 'error');
        }

        // Terminate worker to free resources
        worker.terminate();
    };

    worker.onerror = (error) => {
        console.error('Worker error:', error);
        const errorMessage = error.message || 'Bilinmeyen hata';
        const errorData = {
            id: tempId,
            fileName: file.name,
            title: '-',
            period: '-',
            dateRange: '-',
            processingTime: 'Hata',
            status: 'error'
        };
        addOrUpdateGridRow(errorData);
        showNotification(`${file.name} işlem hatası: ${errorMessage}`, 'error');
        worker.terminate();
    };

    // Start processing
    worker.postMessage({
        file: file,
        fileName: file.name
    });
}

// Add or Update Grid Row
function addOrUpdateGridRow(data) {
    // Check if row already exists
    const existingIndex = parsedFiles.findIndex(f => f.id === data.id);

    if (existingIndex !== -1) {
        parsedFiles[existingIndex] = data;
    } else {
        parsedFiles.push(data);
    }

    renderGrid();
}

// Render Grid
function renderGrid() {
    if (parsedFiles.length === 0) {
        resultsTable.classList.remove('active');
        emptyState.classList.remove('hidden');
        return;
    }

    resultsTable.classList.add('active');
    emptyState.classList.add('hidden');

    // Clear current body
    resultsBody.innerHTML = '';

    // Render all files (newest first)
    [...parsedFiles].reverse().forEach((file, index) => {
        const row = document.createElement('tr');

        // Add status class
        if (file.status === 'processing') row.classList.add('processing-row');
        if (file.status === 'error') row.classList.add('error-row');

        const timeDisplay = file.status === 'processing'
            ? '<span class="badge-processing">⏳ İşleniyor...</span>'
            : (file.status === 'error' ? '❌ Hata' : `<span class="badge-time">${file.processingTime}ms</span>`);

        row.innerHTML = `
            <td><strong>${escapeHtml(file.fileName)}</strong></td>
            <td>${escapeHtml(file.title || '')}</td>
            <td>${escapeHtml(file.period || '')}</td>
            <td>${escapeHtml(file.dateRange || '')}</td>
            <td>${timeDisplay}</td>
        `;

        resultsBody.appendChild(row);
    });
}

// Notifications
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;

    document.body.appendChild(notification);

    // Trigger animation
    setTimeout(() => notification.classList.add('show'), 10);

    // Remove after 3 seconds
    setTimeout(() => {
        notification.classList.remove('show');
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// Helper to escape HTML
function escapeHtml(text) {
    if (text === null || text === undefined) return '';
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
