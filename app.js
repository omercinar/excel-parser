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

// Web Worker code as a string (inline worker to avoid CORS issues with file://)
const workerCode = `
// Import SheetJS library
importScripts('https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js');

// Listen for messages from the main thread
self.addEventListener('message', async (event) => {
    const { file, fileName } = event.data;

    try {
        // Read the file as ArrayBuffer
        const arrayBuffer = await file.arrayBuffer();
        
        // Parse ONLY the first 5 rows for maximum speed
        const workbook = XLSX.read(arrayBuffer, { 
            type: 'array',
            sheetRows: 5
        });
        
        // Get the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Helper to get cell value safely
        const getCellValue = (cell) => {
            if (!cell) return '';
            if (cell.t === 'd' || cell.t === 'n') {
                return cell.w || XLSX.SSF.format('dd/MM/yyyy', cell.v) || String(cell.v || '');
            }
            return String(cell.v || cell.w || '').trim();
        };

        const title = getCellValue(worksheet['A2']);
        const period = getCellValue(worksheet['B3']);
        const dateRange = getCellValue(worksheet['B4']);
        
        // Send success message back to main thread
        self.postMessage({
            success: true,
            data: {
                fileName: fileName,
                title: title,
                period: period,
                dateRange: dateRange
            }
        });
        
    } catch (error) {
        self.postMessage({
            success: false,
            error: error.message || 'Excel dosyası işlenirken bir hata oluştu'
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
            file.name.endsWith('.xlsx') ||
            file.name.endsWith('.xls');

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
        showNotification(`${file.name} işlem hatası`, 'error');
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
