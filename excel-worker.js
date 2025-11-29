// Import SheetJS library
importScripts('https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js');

// Listen for messages from the main thread
self.addEventListener('message', async (event) => {
    const { file, fileName } = event.data;

    try {
        // Read the file as ArrayBuffer
        const arrayBuffer = await file.arrayBuffer();
        
        // Parse the Excel file
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        // Get the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Get cell A4
        const cellA4 = worksheet['A4'];
        
        if (!cellA4) {
            throw new Error('A4 hücresi bulunamadı');
        }
        
        // Get the value from A4
        let cellValue = cellA4.v || cellA4.w || '';
        
        // If the cell is a date type, convert it to string
        if (cellA4.t === 'd' || cellA4.t === 'n') {
            // Try to get formatted value
            cellValue = cellA4.w || XLSX.SSF.format('dd/MM/yyyy', cellA4.v);
        }
        
        cellValue = String(cellValue).trim();
        
        // Parse the date range format: dd/MM/yyyy-dd/MM/yyyy
        const dateRangePattern = /(\d{2}\/\d{2}\/\d{4})\s*-\s*(\d{2}\/\d{2}\/\d{4})/;
        const match = cellValue.match(dateRangePattern);
        
        if (!match) {
            throw new Error(`A4 hücresindeki format geçersiz. Beklenen format: dd/MM/yyyy-dd/MM/yyyy. Bulunan: "${cellValue}"`);
        }
        
        const startDate = match[1];
        const endDate = match[2];
        const dateRange = `${startDate} - ${endDate}`;
        
        // Validate date format
        const datePattern = /^\d{2}\/\d{2}\/\d{4}$/;
        if (!datePattern.test(startDate) || !datePattern.test(endDate)) {
            throw new Error('Tarih formatı geçersiz. Format: dd/MM/yyyy olmalıdır');
        }
        
        // Send success message back to main thread
        self.postMessage({
            success: true,
            data: {
                fileName: fileName,
                dateRange: dateRange,
                startDate: startDate,
                endDate: endDate
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
