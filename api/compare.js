
import Busboy from "busboy";
import ExcelJS from "exceljs";
import { Readable } from "stream";
import csvParser from "csv-parser";

export const config = {
  api: {
    bodyParser: false,
  },
};

function parseCSV(buffer) {
  return new Promise((resolve, reject) => {
    const rows = [];
    Readable.from(buffer.toString())
      .pipe(csvParser())
      .on("data", (row) => rows.push(row))
      .on("end", () => resolve(rows))
      .on("error", reject);
  });
}

function formatName(row, col1, col2) {
  const first = (row.getCell(col1).value || "").toString().trim().toLowerCase();
  const last = (row.getCell(col2).value || "").toString().trim().toLowerCase();
  return `${first} ${last}`.trim();
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).send("Method not allowed");
  }

  const buffers = {};
  const busboy = Busboy({ headers: req.headers });

  const parseForm = () =>
    new Promise((resolve, reject) => {
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

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffers["mainFile"]);

  const doors = ["470", "471", "473", "474", "476", "477"];

  for (const door of doors) {
    const sheetName = `Door ${door}`;
    const csvBuffer = buffers[`door${door}`];
    const sheet = workbook.getWorksheet(sheetName);
    if (!csvBuffer || !sheet) continue;

    // Step 1: Remove columns A–C from row 4 onward and shift left
    for (let i = 4; i <= sheet.rowCount; i++) {
      sheet.spliceColumns(1, 3);
    }

    // Step 2: Remove leftover highlights and extra values like "approved"
    for (let i = 4; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      row.eachCell((cell, colNumber) => {
        if (typeof cell.value === "string" && cell.value.toLowerCase().includes("approved")) {
          cell.value = null;
        }
        if (cell.fill) {
          cell.fill = null;
        }
      });
    }

    // Step 3: Remove empty rows in A/B
    for (let i = sheet.rowCount; i >= 4; i--) {
      const row = sheet.getRow(i);
      if (!row.getCell(1).value && !row.getCell(2).value) {
        sheet.spliceRows(i, 1);
      }
    }

    // Step 4: Parse CSV and paste into columns D & E from row 4
    const csvRows = await parseCSV(csvBuffer);
    const csvData = csvRows.map((r) => [Object.values(r)[0], Object.values(r)[1]]);

    let insertRow = 4;
    for (const [last, first] of csvData) {
      const row = sheet.getRow(insertRow++);
      row.getCell(4).value = last;
      row.getCell(5).value = first;
    }

    // Step 5: Align rows by inserting blanks when names don't match
    let i = 4;
    while (i <= sheet.rowCount) {
      const oldName = formatName(sheet.getRow(i), 1, 2);
      const newName = formatName(sheet.getRow(i), 4, 5);

      if (oldName && newName && oldName === newName) {
        i++;
        continue;
      }

      if (!oldName && newName) {
        // New name added
        sheet.getRow(i).getCell(4).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
        sheet.getRow(i).getCell(5).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
        i++;
        continue;
      }

      if (oldName && !newName) {
        // Access removed
        sheet.getRow(i).getCell(1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" },
        };
        sheet.getRow(i).getCell(2).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF0000" },
        };
        i++;
        continue;
      }

      const rightNames = formatName(sheet.getRow(i + 1), 4, 5);
      const leftNames = formatName(sheet.getRow(i + 1), 1, 2);

      if (oldName && newName && oldName !== newName) {
        if (newName === formatName(sheet.getRow(i + 1), 1, 2)) {
          // Insert blank row in A/B
          sheet.spliceRows(i, 0, []);
        } else if (oldName === formatName(sheet.getRow(i + 1), 4, 5)) {
          // Insert blank row in D/E
          sheet.spliceRows(i, 0, []);
        } else {
          i++;
        }
      } else {
        i++;
      }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  res.setHeader("Content-Disposition", "attachment; filename=highlighted.xlsx");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
}
