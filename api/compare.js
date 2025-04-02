// api/compare.js
import Busboy from "busboy";
import ExcelJS from "exceljs";
import { Readable } from "stream";
import csv from "csv-parser";

export const config = {
  api: {
    bodyParser: false,
  },
};

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

      busboy.on("finish", () => resolve());
      busboy.on("error", reject);
      req.pipe(busboy);
    });

  await parseForm();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffers.mainFile);

  const doors = ["470", "471", "473", "474", "476", "477"];
  for (const door of doors) {
    const sheet = workbook.getWorksheet(`Door ${door}`);
    if (!sheet) continue;

    // Step 6: Delete columns A–C from row 4 down
    for (let rowIndex = 4; rowIndex <= sheet.rowCount; rowIndex++) {
      sheet.spliceColumns(1, 3, [], []);
    }

    // Step 8: Remove blank rows where new Column A or B is empty
    for (let i = sheet.rowCount; i >= 4; i--) {
      const row = sheet.getRow(i);
      if (!row.getCell(1).value && !row.getCell(2).value) {
        sheet.spliceRows(i, 1);
      }
    }

    // Step 9: Parse the uploaded CSV for this door
    const csvRows = [];
    await new Promise((resolve, reject) => {
      Readable.from(buffers[`door${door}`])
        .pipe(csv())
        .on("data", (data) => {
          const keys = Object.keys(data);
          if (keys.length >= 2) {
            csvRows.push([data[keys[0]], data[keys[1]]]);
          }
        })
        .on("end", resolve)
        .on("error", reject);
    });

    // Step 10–16: Copy CSV names into column D/E and align
    let rowPointer = 4;
    for (const [lastName, firstName] of csvRows) {
      const targetRow = sheet.getRow(rowPointer);
      targetRow.getCell(4).value = lastName;
      targetRow.getCell(5).value = firstName;
      rowPointer++;
    }

    // Step 15–16: Highlight new names and align rows
    for (let i = 4; i < rowPointer; i++) {
      const row = sheet.getRow(i);
      const oldName = `${row.getCell(1).value || ""}${row.getCell(2).value || ""}`.toLowerCase().trim();
      const newName = `${row.getCell(4).value || ""}${row.getCell(5).value || ""}`.toLowerCase().trim();

      if (!oldName && newName) {
        row.getCell(4).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF00" },
        };
        row.getCell(5).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF00" },
        };
      } else if (oldName && !newName) {
        row.getCell(4).value = "";
        row.getCell(5).value = "";
      }
    }
  }

  const outputBuffer = await workbook.xlsx.writeBuffer();

  res.setHeader("Content-Disposition", 'attachment; filename="highlighted.xlsx"');
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(outputBuffer);
}
