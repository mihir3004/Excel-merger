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
  function getCellFillHex(cell) {
    if (
      cell.fill &&
      cell.fill.type === "pattern" &&
      cell.fill.fgColor &&
      cell.fill.fgColor.argb
    ) {
      // ARGB is like 'FFFFC000' — drop the first two 'FF' (alpha)
      return `#${cell.fill.fgColor.argb.slice(2)}`;
    }
  
    return null; // No color or unsupported format
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
        ); 
      const dataStartRow = headerRowCount + 1; 

      if (index === 0) {
        (sheet.model.merges || []).forEach((merge) => {
          worksheet.mergeCells(merge);
        });
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

      rows.forEach((row, rowIndex) => {
        const sourceRow = sheet.getRow(rowIndex + dataStartRow);
       
        const isEmptyRow = sourceRow.values.every(
          (cell) => cell === null || cell === ""
        );

        if (!isEmptyRow) {
          const newRow = worksheet.addRow([]);
          sourceRow.eachCell((cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);
            newCell.value = cell.formula ? cell.result || null : cell.value;
            newCell.style = { ...cell.style };
            const hexColor = getCellFillHex(cell);
    if (hexColor) {
      console.log(`Cell at ${cell.address} has fill color: ${hexColor}`);
    }

          });
          if (newRow.getCell(1).value) {
            newRow.getCell(1).value = serialNumber++;
          }
        }
      });
    }
    worksheet.columns.forEach((col) => {
      col.width = 25;
    });
  }
  // const yellowFilteredSheet = workbook.addWorksheet("Vehicle Release List");
  const sourceSheet = workbook.getWorksheet(1);
  // // Copy header from first sheet (3 rows like Sheet1)
  // const headerRow = yellowFilteredSheet.getRow(2); // Adjust row number if needed
  // headerRow.getCell(1).value = "Vehicle Release";
  // headerRow.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
  // headerRow.getCell(1).font = { bold: true, size: 18 };
  // yellowFilteredSheet.mergeCells(2, 1, 2, sourceSheet.columnCount); // Merge from column 1 to last
  
  // // Add headers with bold, centered, and bordered styles
  // for (let j = 3; j <= 3; j++) {
  //   const sourceRow = sourceSheet.getRow(j);
  //   const newRow = yellowFilteredSheet.getRow(j);
  
  //   sourceRow.eachCell((cell, colNumber) => {
  //     const newCell = newRow.getCell(colNumber);
  //     newCell.value = cell.value;
  
  //     newCell.alignment = { horizontal: "center", vertical: "middle" };
  //     newCell.font = { bold: true,size:18,...newCell.font };
  
  //     newCell.border = {
  //       top: { style: "thin" },
  //       left: { style: "thin" },
  //       bottom: { style: "thin" },
  //       right: { style: "thin" },
  //     };
  //   });
  // }
  
  
  // let redRowSerial = 1;
  // let destRowIndex = 4; // Start writing data from row 4
  // for (let i = 1; i <= sheetCount; i++) {
  //   const mergedSheet = workbook.getWorksheet(i);
  //   mergedSheet.eachRow((row, rowNumber) => {
  //     if (rowNumber <= 3) return; // Skip headers
  
  //     const targetCell = row.getCell(10);
  //     const fillHex = getCellFillHex(targetCell);
  
  //     if (fillHex === "#FFFF00") {
  //       const newRow = yellowFilteredSheet.getRow(destRowIndex++);
  //       row.eachCell((cell, colNumber) => {
  //         const newCell = newRow.getCell(colNumber);
  //         newCell.value = cell.value;
  //         newCell.style = { ...cell.style };
  //       });
  
  
  //       // Assign serial number to column 1
  //       newRow.getCell(1).value = redRowSerial++;
  //     }
  //   });
    
    
  // }
  // yellowFilteredSheet.columns.forEach((col) => {
  //   col.width = 25; // You can change the width value as needed
  // });
  const orangeFilteredSheet = workbook.addWorksheet("Vehicle Release List");
  const headerRow1 = orangeFilteredSheet.getRow(2); // Adjust row number if needed

  headerRow1.getCell(1).value = "Vehicle Release";
  headerRow1.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
  headerRow1.getCell(1).font = { bold: true, size: 18 };
  orangeFilteredSheet.mergeCells(2, 1, 2, sourceSheet.columnCount); // Merge from column 1 to last
  
  // Add headers with bold, centered, and bordered styles
  for (let j = 3; j <= 3; j++) {
    const sourceRow = sourceSheet.getRow(j);
    const newRow = orangeFilteredSheet.getRow(j);
  
    sourceRow.eachCell((cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
      newCell.value = cell.value;
  
      newCell.alignment = { horizontal: "center", vertical: "middle" };
      newCell.font = { bold: true,size:18,...newCell.font };
  
      newCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
  }
  
  
   redRowSerial = 1;
  destRowIndex = 4; // Start writing data from row 4
  for (let i = 1; i <= sheetCount; i++) {
    const mergedSheet = workbook.getWorksheet(i);
    mergedSheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 3) return; // Skip headers
  
      const targetCell = row.getCell(11);
      const fillHex = getCellFillHex(targetCell);

      if (fillHex === "#FFC000") {
        const newRow = orangeFilteredSheet.getRow(destRowIndex++);
        row.eachCell((cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          newCell.value = cell.value;
          newCell.style = { ...cell.style };
        });
  
        // Assign serial number to column 1
        newRow.getCell(1).value = redRowSerial++;
      }
    });
    
    
  }
  orangeFilteredSheet.columns.forEach((col) => {
    col.width = 25; // You can change the width value as needed
  });






  const yellowFilteredSheet = workbook.addWorksheet("Vehicle Detained List");
  const headerRow = yellowFilteredSheet.getRow(2); // Adjust row number if needed

  headerRow.getCell(1).value = "Vehicle Detained";
  headerRow.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
  headerRow.getCell(1).font = { bold: true, size: 18 };
  yellowFilteredSheet.mergeCells(2, 1, 2, sourceSheet.columnCount); // Merge from column 1 to last
  
  // Add headers with bold, centered, and bordered styles
  for (let j = 3; j <= 3; j++) {
    const sourceRow = sourceSheet.getRow(j);
    const newRow = yellowFilteredSheet.getRow(j);
  
    sourceRow.eachCell((cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
      newCell.value = cell.value;
  
      newCell.alignment = { horizontal: "center", vertical: "middle" };
      newCell.font = { bold: true,size:18,...newCell.font };
  
      newCell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
  }
  
  
   redRowSerial = 1;
  destRowIndex = 4; // Start writing data from row 4
  for (let i = 1; i <= sheetCount; i++) {
    const mergedSheet = workbook.getWorksheet(i);
    mergedSheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 3) return; // Skip headers
  
      const targetCell = row.getCell(10);
      const fillHex = getCellFillHex(targetCell);

      if (fillHex === "#FFFF00") {
        const newRow = yellowFilteredSheet.getRow(destRowIndex++);
        row.eachCell((cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          newCell.value = cell.value;
          newCell.style = { ...cell.style };
        });
  
        // Assign serial number to column 1
        newRow.getCell(1).value = redRowSerial++;
      }
    });
    
    
  }
  yellowFilteredSheet.columns.forEach((col) => {
    col.width = 25; // You can change the width value as needed
  });





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
