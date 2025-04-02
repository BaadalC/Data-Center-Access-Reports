import Busboy from "busboy";
import ExcelJS from "exceljs";
import { Readable } from "stream";
import * as XLSX from "xlsx";

export const config = {
  api: {
    bodyParser: false,
  },
};

const yellowFill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFF00" },
};

const redFill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF0000" },
};

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const buffers = {};
  const busboy = Busboy({ headers: req.headers });

  const parseForm = new Promise((resolve, reject) => {
    busboy.on("file", (fieldname, file) => {
      const chunks = [];
      file.on("data", (data) => chunks.push(data));
      file.on("end", () => {
        buffers[fieldname] = Buffer.concat(chunks);
      });
    });
    busboy.on("finish", resolve);
    busboy.on("error", reject);
    req.pipe(busboy);
  });

  await parseForm();

  const mainWorkbook = new ExcelJS.Workbook();
  await mainWorkbook.xlsx.load(buffers["mainFile"]);
  const newWorkbook = new ExcelJS.Workbook();

  for (const field in buffers) {
    if (field === "mainFile") continue;

    const doorNumber = field.replace("door", "");
    const tabName = `Door ${doorNumber}`;
    const sheet = mainWorkbook.getWorksheet(tabName);
    if (!sheet) continue;

    // Delete cols A-C from row 4 down
    for (let row = 4; row <= sheet.rowCount; row++) {
      for (let col = 1; col <= 3; col++) {
        sheet.getCell(row, col).value = null;
      }
    }

    // Shift D/E into A/B
    for (let row = 4; row <= sheet.rowCount; row++) {
      sheet.getCell(row, 1).value = sheet.getCell(row, 4).value;
      sheet.getCell(row, 2).value = sheet.getCell(row, 5).value;
      sheet.getCell(row, 4).value = null;
      sheet.getCell(row, 5).value = null;
    }

    // Delete rows with blank A/B
    let row = 4;
    while (row <= sheet.rowCount) {
      const last = sheet.getCell(row, 1).value;
      const first = sheet.getCell(row, 2).value;
      if (!last && !first) {
        sheet.spliceRows(row, 1);
      } else {
        row++;
      }
    }

    // Parse new CSV
    const csv = XLSX.read(buffers[field], { type: "buffer" });
    const csvSheet = csv.Sheets[csv.SheetNames[0]];
    const csvJson = XLSX.utils.sheet_to_json(csvSheet, { header: 1 });
    const csvNames = csvJson.map((row) => [row[0], row[1]]).filter(r => r[0] || r[1]);

    // Paste into D/E starting row 4
    csvNames.forEach(([last, first], i) => {
      sheet.getCell(i + 4, 4).value = last;
      sheet.getCell(i + 4, 5).value = first;
    });

    // Align and highlight
    const maxLen = Math.max(sheet.rowCount - 3, csvNames.length);
    for (let i = 0; i < maxLen; i++) {
      const rowNum = i + 4;
      const oldLast = sheet.getCell(rowNum, 1).value;
      const oldFirst = sheet.getCell(rowNum, 2).value;
      const newLast = sheet.getCell(rowNum, 4).value;
      const newFirst = sheet.getCell(rowNum, 5).value;

      const oldFull = `${oldLast || ""},${oldFirst || ""}`;
      const newFull = `${newLast || ""},${newFirst || ""}`;

      if (oldFull !== newFull) {
        if (newLast || newFirst) {
          sheet.getCell(rowNum, 4).fill = yellowFill;
          sheet.getCell(rowNum, 5).fill = yellowFill;
        }
        if (oldLast || oldFirst) {
          sheet.getCell(rowNum, 1).fill = redFill;
          sheet.getCell(rowNum, 2).fill = redFill;
        }
      }
    }

    // Add updated sheet to output workbook
    const newSheet = newWorkbook.addWorksheet(tabName);
    sheet.eachRow({ includeEmpty: true }, (row, rowNum) => {
      const newRow = newSheet.getRow(rowNum);
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        newRow.getCell(colNum).value = cell.value;
        if (cell.fill) newRow.getCell(colNum).fill = cell.fill;
      });
      newRow.commit();
    });
  }

  const buffer = await newWorkbook.xlsx.writeBuffer();
  res.setHeader("Content-Disposition", "attachment; filename=highlighted.xlsx");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
}
