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

  // Process door1 - door6
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

    // Step 3: Parse reference file
    const refBuffer = refFile.buffer;
    const type = await fileTypeFromBuffer(refBuffer);

    let referenceNames = [];
    if (type?.ext === "csv") {
      const parsed = parse(refBuffer.toString(), {
        skip_empty_lines: true,
      });
      referenceNames = parsed.map((row) => ({
        last: row[0]?.trim(),
        first: row[1]?.trim(),
        full: `${row[0]?.trim()} ${row[1]?.trim()}`.trim(),
      }));
    } else if (type?.ext === "xlsx") {
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
    } else {
      continue; // skip unsupported files
    }

    // Step 4: Insert into columns D & E, from row 4
    for (let r = 0; r < referenceNames.length; r++) {
      const name = referenceNames[r];
      const rowIndex = r + 4;
      const row = tab.getRow(rowIndex);
      row.getCell(4).value = name.last;
      row.getCell(5).value = name.first;
    }

    // Step 5: Compare A&B with D&E
    const oldNames = [];
    const newNames = [];

    for (let rowNum = 4; rowNum <= tab.rowCount; rowNum++) {
      const row = tab.getRow(rowNum);
      const oldFull = `${row.getCell(1).value || ""} ${row.getCell(2).value || ""}`.trim();
      const newFull = `${row.getCell(4).value || ""} ${row.getCell(5).value || ""}`.trim();

      if (oldFull && !referenceNames.find((n) => n.full === oldFull)) {
        // Removed
        row.getCell(1).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0000" },
        };
        row.getCell(2).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0000" },
        };
      }

      if (newFull && !tab.getColumn(1).values.includes(newFull)) {
        // Added
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
      }
    }
  }

  // Send updated file
  const buffer = await mainWorkbook.xlsx.writeBuffer();
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="highlighted-report.xlsx"');
  res.send(Buffer.from(buffer));
}
