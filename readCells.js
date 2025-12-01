const fs = require("fs");
const unzip = require("unzipper");
const sax = require("sax");

// Okunacak hücreler
const TARGET_CELLS = ["A2", "B3", "B4"];

// Program argümanından dosya adı
const file = process.argv[2];
if (!file) {
    console.error("Kullanım: node readC6.js <excel.xlsx>");
    process.exit(1);
}

async function readCells(xlsxPath) {
    const startTime = Date.now();

    // ZIP içinden sheet1.xml dosyasını aç
    const directory = await unzip.Open.file(xlsxPath);
    const sheetEntry = directory.files.find(f => f.path === "xl/worksheets/sheet1.xml");

    if (!sheetEntry) {
        console.error("sheet1.xml bulunamadı");
        return null;
    }

    const stream = sheetEntry.stream();
    const saxStream = sax.createStream(true);

    const results = {
        "A2": null,
        "B3": null,
        "B4": null
    };

    let currentCell = null;
    let readingValue = false;

    saxStream.on("opentag", (node) => {
        if (node.name === "c") {
            const cellRef = node.attributes.r;
            if (TARGET_CELLS.includes(cellRef)) {
                currentCell = cellRef;
            }
        }
        // Değer okuma: <v> veya <t> (inline string)
        if (currentCell && (node.name === "v" || node.name === "t")) {
            readingValue = true;
        }
    });

    saxStream.on("text", (text) => {
        if (readingValue && currentCell) {
            // Mevcut değerin üzerine ekle (büyük metinler parça parça gelebilir)
            results[currentCell] = (results[currentCell] || "") + text;
        }
    });

    saxStream.on("closetag", (name) => {
        if (name === "v" || name === "t") {
            readingValue = false;
        }
        if (name === "c") {
            currentCell = null;

            // Check if all targets are found
            const allFound = TARGET_CELLS.every(cell => results[cell] !== null);
            if (allFound) {
                // Stop processing
                stream.unpipe(saxStream);
                saxStream.end();
                stream.destroy(); // Stop reading the file
            }
        }
    });

    return new Promise((resolve) => {
        let resolved = false;

        const finish = () => {
            if (!resolved) {
                resolved = true;
                const duration = Date.now() - startTime;
                resolve({ results, duration });
            }
        };

        saxStream.on("end", finish);
        // In case unpipe/destroy doesn't trigger end immediately or triggers error
        saxStream.on("error", (e) => {
            // Ignore errors if we are done
            if (!resolved) console.warn("Stream error:", e.message);
            finish();
        });
        stream.on("close", finish);

        stream.pipe(saxStream);
    });
}

(async () => {
    try {
        const output = await readCells(file);

        if (output) {
            console.log(`A2 -> ${output.results["A2"] || "Bulunamadı"}`);
            console.log(`B3 -> ${output.results["B3"] || "Bulunamadı"}`);
            console.log(`B4 -> ${output.results["B4"] || "Bulunamadı"}`);
            console.log(`İşlem Süresi: ${output.duration}ms`);
        }
    } catch (error) {
        console.error("Hata:", error.message);
    }
})();
