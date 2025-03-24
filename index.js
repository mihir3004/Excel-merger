const express = require("express");
const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");
const ExcelJS = require("exceljs");

const app = express();
const PORT = 5000;
const excelFolderPath = path.join(__dirname, "public");
const mergeExcelData = async () => {
  const files = fs
    .readdirSync(excelFolderPath)
    .filter((file) => file.endsWith(".xlsx"));

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
    let serialNumber = 1;

    for (const [index, file] of excelData.entries()) {
      const sheet = file.data.worksheets[i];
      const headerRowCount = i === 0 ? 3 : 1;
      const rows = sheet
        .getSheetValues()
        .slice(headerRowCount)
        .filter(
          (row) =>
            row && row.every((cell) => cell !== null && cell.value !== "")
        ); // Remove empty first index
      const dataStartRow = headerRowCount + 1; // Data starts after header

      if (index === 0) {
        (sheet.model.merges || []).forEach((merge) => {
          worksheet.mergeCells(merge);
        });
        // Copy first 3 header rows with styles
        for (let j = 1; j <= headerRowCount; j++) {
          const sourceRow = sheet.getRow(j);
          const newRow = worksheet.getRow(j);
          sourceRow.eachCell((cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);
            newCell.value = cell.formula ? cell.result || null : cell.value;
            newCell.style = { ...cell.style };
          });
        }
      }

      // Extract data rows (skip headers)
      //   let dataRows = rows.slice(dataStartRow);

      //   if (i == 1) {
      //     dataRows = rows.slice(1);

      //     // console.log(rows);
      //   }
      rows.forEach((row, rowIndex) => {
        const sourceRow = sheet.getRow(rowIndex + dataStartRow);
        if (i == 0 && index == 0) {
          console.log(sourceRow.values);
        }
        const isEmptyRow = sourceRow.values.every(
          (cell) => cell === null || cell === ""
        );

        if (!isEmptyRow) {
          const newRow = worksheet.addRow([]);
          sourceRow.eachCell((cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);
            newCell.value = cell.formula ? cell.result || null : cell.value;
            newCell.style = { ...cell.style };
          });
          if (newRow.getCell(1).value) {
            newRow.getCell(1).value = serialNumber++;
          }
        }
      });
    }

    // Set default column width
    worksheet.columns.forEach((col) => {
      col.width = 25;
    });
  }

  const outputPath = path.join(
    __dirname,
    "output",
    `merged-${Date.now()}.xlsx`
  );
  await workbook.xlsx.writeFile(outputPath);
  console.log(`✅ Merged file saved at: ${outputPath}`);
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
