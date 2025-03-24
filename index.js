const express = require("express");
const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");
const ExcelJS = require("exceljs");

const app = express();
const PORT = 5000;
const excelFolderPath = path.join(__dirname, "public");
const mergeExcelData = async () => {
    const files = fs.readdirSync(excelFolderPath).filter(file => file.endsWith(".xlsx"));

    if (files.length !== 3) {
        console.log("Error: Exactly three Excel files are required.");
        return null;
    }

    const workbook = new ExcelJS.Workbook();
    const excelData = await Promise.all(
        files.map(async (file) => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(path.join(excelFolderPath, file));
            return { name: file, data: wb };
        })
    );

    const sheetCount = excelData[0].data.worksheets.length;

    for (const file of excelData) {
        if (file.data.worksheets.length !== sheetCount) {
            console.log(`Error: Mismatched sheet count in file ${file.name}`);
            return null;
        }
    }

    for (let i = 0; i < sheetCount; i++) {
        const worksheet = workbook.addWorksheet(`Sheet${i + 1}`);
        let mergedData = [];
        let serialNumber = 1;

        for (const [index, file] of excelData.entries()) {
            const sheet = file.data.worksheets[i];
            const rows = sheet.getSheetValues().slice(1); // Remove empty first index

            // Identify header row index (first non-empty row)
            const headerRowIndex = rows.findIndex(row => row && row.some(cell => cell !== null && cell !== ""));

            if (headerRowIndex === -1) continue;

            if (index === 0) {
                // Copy header rows including styles
                for (let j = 0; j < 3; j++) {
                    const rowNumber = headerRowIndex + j;
                    const sourceRow = sheet.getRow(rowNumber);
                    const newRow = worksheet.getRow(mergedData.length + 1);
                    sourceRow.eachCell((cell, colNumber) => {
                        const newCell = newRow.getCell(colNumber);
                    
                        // Instead of copying the formula, get only the value
                        if (cell.formula) {
                            newCell.value = cell.result || null; // Use calculated value
                        } else {
                            newCell.value = cell.value;
                        }
                    
                        // Copy cell styles if needed
                        newCell.style = { ...cell.style };
                    });
                    
                    
                    
                    
                    mergedData.push([]);
                }
            }

            // Extract data rows (skipping headers)
            const dataRows = rows.slice(headerRowIndex + 3).filter(row => row && row.some(cell => cell !== null && cell !== ""));

            // Update Sr. Number column while keeping styles
            dataRows.forEach((row, rowIndex) => {
                if (row.length > 0) row[1] = serialNumber++; // Assuming Sr. No is in the first column

                const sourceRow = sheet.getRow(headerRowIndex + 3 + rowIndex);
                const newRow = worksheet.getRow(mergedData.length + 1);

                sourceRow.eachCell((cell, colNumber) => {
                    const newCell = newRow.getCell(colNumber);
                
                    // Instead of copying the formula, get only the value
                    if (cell.formula) {
                        newCell.value = cell.result || null; // Use calculated value
                    } else {
                        newCell.value = cell.value;
                    }
                
                    // Copy cell styles if needed
                    newCell.style = { ...cell.style };
                });
                

                mergedData.push([]);
            });
        }

        worksheet.columns.forEach(col => {
            col.width = 15; // Set a reasonable default width
        });
    }

    const outputPath = path.join(__dirname, "public", "merged.xlsx");
    await workbook.xlsx.writeFile(outputPath);
    return outputPath;
};





app.get("/merge-excel", async (req, res) => {
    const mergedFilePath = await mergeExcelData();

    if (!mergedFilePath) {
        return res.status(400).json({ error: "Error merging Excel files" });
    }

    res.download(mergedFilePath);
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
