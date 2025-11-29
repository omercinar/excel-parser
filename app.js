// Initialize Web Worker
let worker = null;
let parsedFiles = [];
let processingStartTime = 0;
let timerInterval = null;

// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const uploadButton = document.getElementById('uploadButton');
const fileInput = document.getElementById('fileInput');
const loadingIndicator = document.getElementById('loadingIndicator');
const resultsSection = document.getElementById('resultsSection');
const resultsBody = document.getElementById('resultsBody');
const resultsTable = document.querySelector('.results-table');
const emptyState = document.getElementById('emptyState');

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
        // This is much faster than parsing the entire file
        const workbook = XLSX.read(arrayBuffer, { 
            type: 'array',
            sheetRows: 5  // Only parse first 5 rows (we need A2, B3, B4)
        });
        
        // Get the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Get cell A2 (Ünvan)
        const cellA2 = worksheet['A2'];
        let title = '';
        
        if (cellA2) {
            // Get the value from A2
            if (cellA2.t === 'd' || cellA2.t === 'n') {
                title = cellA2.w || XLSX.SSF.format('dd/MM/yyyy', cellA2.v) || String(cellA2.v || '');
            } else {
                title = String(cellA2.v || cellA2.w || '').trim();
            }
        }
        
        // Get cell B3 (Dönem)
        const cellB3 = worksheet['B3'];
        let period = '';
        
        if (cellB3) {
            // Get the value from B3
            if (cellB3.t === 'd' || cellB3.t === 'n') {
                period = cellB3.w || XLSX.SSF.format('dd/MM/yyyy', cellB3.v) || String(cellB3.v || '');
            } else {
                period = String(cellB3.v || cellB3.w || '').trim();
            }
        }
        
        // Get cell B4 (Tarih Aralığı)
        const cellB4 = worksheet['B4'];
        let dateRange = '';
        
        if (cellB4) {
            // Get the value from B4
            if (cellB4.t === 'd' || cellB4.t === 'n') {
                dateRange = cellB4.w || XLSX.SSF.format('dd/MM/yyyy', cellB4.v) || String(cellB4.v || '');
            } else {
                dateRange = String(cellB4.v || cellB4.w || '').trim();
            }
        }
        
        // Send success message back to main thread
        // No validation - just send whatever is in the cells
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
        // Send error message back to main thread
        self.postMessage({
            success: false,
            error: error.message || 'Excel dosyası işlenirken bir hata oluştu'
        });
    }
});
`;

// Initialize Web Worker
function initWorker() {
    if (!worker) {
        // Create a Blob from the worker code
        const blob = new Blob([workerCode], { type: 'application/javascript' });
        const workerUrl = URL.createObjectURL(blob);

        // Create worker from Blob URL
        worker = new Worker(workerUrl);

        worker.addEventListener('message', (event) => {
            const processingTime = Date.now() - processingStartTime;
            hideLoading();

            if (event.data.success) {
                // Add processing time to data
                const resultData = {
                    ...event.data.data,
                    processingTime: processingTime
                };
                addResultToGrid(resultData);
                showSuccessNotification(`Dosya başarıyla işlendi! (${processingTime}ms)`);
            } else {
                showErrorNotification(event.data.error);
            }
        });

        worker.addEventListener('error', (error) => {
            hideLoading();
            showErrorNotification('Web Worker hatası: ' + error.message);
        });
    }
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
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
    // Reset input so the same file can be selected again
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

    const file = e.dataTransfer.files[0];
    if (file) {
        handleFile(file);
    }
});

// Handle File Processing
function handleFile(file) {
    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel' // .xls
    ];

    const isValidType = validTypes.includes(file.type) ||
        file.name.endsWith('.xlsx') ||
        file.name.endsWith('.xls');

    if (!isValidType) {
        showErrorNotification('Lütfen geçerli bir Excel dosyası seçin (.xlsx veya .xls)');
        return;
    }

    // Show loading
    showLoading();

    // Start timer
    processingStartTime = Date.now();

    // Initialize worker if not already done
    initWorker();

    // Send file to worker
    worker.postMessage({
        file: file,
        fileName: file.name
    });
}

// Add Result to Grid
function addResultToGrid(data) {
    // Check if file already exists
    const existingIndex = parsedFiles.findIndex(f => f.fileName === data.fileName);

    if (existingIndex !== -1) {
        // Update existing entry
        parsedFiles[existingIndex] = data;
    } else {
        // Add new entry
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

    resultsBody.innerHTML = '';

    parsedFiles.forEach((file, index) => {
        const row = document.createElement('tr');
        row.style.setProperty('--row-index', index);

        row.innerHTML = `
            <td><strong>${escapeHtml(file.fileName)}</strong></td>
            <td>${escapeHtml(file.title || '')}</td>
            <td>${escapeHtml(file.period || '')}</td>
            <td>${escapeHtml(file.dateRange || '')}</td>
            <td><span class="badge-time">${file.processingTime}ms</span></td>
            <td>
                <button class="delete-button" onclick="deleteFile(${index})">
                    🗑️ Sil
                </button>
            </td>
        `;

        resultsBody.appendChild(row);
    });
}

// Delete File from Grid
function deleteFile(index) {
    parsedFiles.splice(index, 1);
    renderGrid();
    showSuccessNotification('Dosya silindi');
}

// Loading State
function showLoading() {
    uploadArea.style.display = 'none';
    loadingIndicator.classList.add('active');

    // Start live timer
    const timerElement = document.getElementById('processingTimer');
    timerInterval = setInterval(() => {
        const elapsed = Date.now() - processingStartTime;
        timerElement.textContent = `${elapsed}ms`;
    }, 10); // Update every 10ms for smooth counting
}

function hideLoading() {
    uploadArea.style.display = 'block';
    loadingIndicator.classList.remove('active');

    // Stop live timer
    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }
}

// Notifications
function showSuccessNotification(message) {
    showNotification(message, 'success');
}

function showErrorNotification(message) {
    showNotification(message, 'error');
}

function showNotification(message, type = 'info') {
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;

    // Add styles
    Object.assign(notification.style, {
        position: 'fixed',
        top: '20px',
        right: '20px',
        padding: '1rem 1.5rem',
        borderRadius: '12px',
        color: 'white',
        fontWeight: '500',
        boxShadow: '0 8px 32px rgba(0, 0, 0, 0.4)',
        zIndex: '1000',
        animation: 'slideInRight 0.3s ease',
        maxWidth: '400px',
        backdropFilter: 'blur(10px)'
    });

    if (type === 'success') {
        notification.style.background = 'linear-gradient(135deg, #4facfe 0%, #00f2fe 100%)';
    } else if (type === 'error') {
        notification.style.background = 'linear-gradient(135deg, #f093fb 0%, #f5576c 100%)';
    } else {
        notification.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
    }

    document.body.appendChild(notification);

    // Add slide in animation
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideInRight {
            from {
                opacity: 0;
                transform: translateX(100px);
            }
            to {
                opacity: 1;
                transform: translateX(0);
            }
        }
    `;
    document.head.appendChild(style);

    // Remove after 3 seconds
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease';
        notification.style.opacity = '0';
        notification.style.transform = 'translateX(100px)';
        setTimeout(() => {
            notification.remove();
            style.remove();
        }, 300);
    }, 3000);
}

// Utility Functions
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Make deleteFile function globally accessible
window.deleteFile = deleteFile;

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    renderGrid();
});
