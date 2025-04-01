import Busboy from "busboy";
import ExcelJS from "exceljs";
import { Readable } from "stream";
import { fileTypeFromBuffer } from "file-type";
import { parse } from "csv-parse/sync";

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

  const parseForm = new Promise((resolve, reject) => {
    busboy.on("file", (fieldname, file, filename) => {
      const chunks = [];
      file.on("data", (data) => chunks.push(data));
      file.on("end", () => {
        buffers[fieldname] = {
          buffer: Buffer.concat(chunks),
          filename: filename || "",
        };
      });
    });

    busboy.on("finish", resolve);
    busboy.on("error", reject);
    req.pipe(busboy);
  });

  await parseForm;

  const mainFile = buffers["main"];
  if (!mainFile) return res.status(400).send("Missing main file");

  const mainWorkbook = new ExcelJS.Workbook();
  await mainWorkbook.xlsx.load(mainFile.buffer);

  for (let i = 1; i <= 6; i++) {
    const field = `door${i}`;
    const refFile = buffers[field];
    if (!refFile) continue;

    const tab = mainWorkbook.getWorksheet(i);
    if (!tab) continue;

    // Step 1: Delete columns A–C from row 4+
    for (let rowNum = tab.rowCount; rowNum >= 4; rowNum--) {
      tab.getRow(rowNum).splice(1, 3);
    }

    // Step 2: Remove blank rows (A & B)
    for (let rowNum = tab.rowCount; rowNum >= 4; rowNum--) {
      const row = tab.getRow(rowNum);
      const a = row.getCell(1).value?.toString().trim();
      const b = row.getCell(2).value?.toString().trim();
      if (!a && !b) tab.spliceRows(rowNum, 1);
    }

    // Step 3: Detect file type
    const refBuffer = refFile.buffer;
    const type = await fileTypeFromBuffer(refBuffer);
    console.log(`Detected file type for ${field}:`, type);

    if (!type || (type.ext !== 'csv' && type.ext !== 'xlsx')) {
      console.warn(`Unsupported file type for ${field}:`, type);
      continue;
    }

    // Step 4: Parse file into name list
    let referenceNames = [];

    if (type.ext === "csv") {
      try {
        const parsed = parse(refBuffer.toString(), {
          skip_empty_lines: true,
        });
        referenceNames = parsed.map((row) => ({
          last: row[0]?.trim(),
          first: row[1]?.trim(),
          full: `${row[0]?.trim()} ${row[1]?.trim()}`.trim(),
        }));
      } catch (err) {
        console.error(`Error parsing CSV for ${field}:`, err);
        continue;
      }
    }

    if (type.ext === "xlsx") {
      try {
        const refWorkbook = new ExcelJS.Workbook();
        await refWorkbook.xlsx.load(refBuffer);
        const sheet = refWorkbook.worksheets[0];
        sheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const last = row.getCell(1).value?.toString().trim();
          const first = row.getCell(2).value?.toString().trim();
          if (last || first) {
            referenceNames.push({ last, first, full: `${last} ${first}`.trim() });
          }
        });
      } catch (err) {
        console.error(`Error parsing XLSX for ${field}:`, err);
        continue;
      }
    }

    // Step 5: Paste into D & E
    for (let r = 0; r < referenceNames.length; r++) {
      const name = referenceNames[r];
      const rowIndex = r + 4;
      const row = tab.getRow(rowIndex);
      row.getCell(4).value = name.last;
      row.getCell(5).value = name.first;
    }

    // Step 6: Compare A/B with D/E
    for (let rowNum = 4; rowNum <= tab.rowCount; rowNum++) {
      const row = tab.getRow(rowNum);
      const oldFull = `${row.getCell(1).value || ""} ${row.getCell(2).value || ""}`.trim();
      const newFull = `${row.getCell(4).value || ""} ${row.getCell(5).value || ""}`.trim();

      const highlight = (cell, color) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: color },
        };
      };

      if (oldFull && !referenceNames.find((n) => n.full === oldFull)) {
        highlight(row.getCell(1), "FF0000"); // Red
        highlight(row.getCell(2), "FF0000");
      }

      if (newFull && oldFull !== newFull && !tab.getColumn(1).values.includes(newFull)) {
        highlight(row.getCell(4), "FFFF00"); // Yellow
        highlight(row.getCell(5), "FFFF00");
      }
    }
  }

  // Send updated file
  const buffer = await mainWorkbook.xlsx.writeBuffer();
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="highlighted-report.xlsx"');
  res.send(Buffer.from(buffer));
}
