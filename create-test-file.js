const fs = require('fs');

// Create a minimal XLSX file structure
// This creates a valid Excel file with the date range in cell B4

const createExcelFile = () => {
    // Minimal XLSX structure (using XML)
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

    const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>`;

    const worksheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="4">
<c r="B4" t="inlineStr">
<is><t>01/01/2024-31/01/2024</t></is>
</c>
</row>
</sheetData>
</worksheet>`;

    // Create directory structure
    const dirs = [
        '_rels',
        'xl',
        'xl/_rels',
        'xl/worksheets'
    ];

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }
    });

    // Write files
    fs.writeFileSync('[Content_Types].xml', contentTypes);
    fs.writeFileSync('_rels/.rels', rels);
    fs.writeFileSync('xl/workbook.xml', workbook);
    fs.writeFileSync('xl/_rels/workbook.xml.rels', workbookRels);
    fs.writeFileSync('xl/worksheets/sheet1.xml', worksheet);

    console.log('Excel structure created. Now zip these files to create test-file.xlsx');
};

createExcelFile();
