// /api/compare.js (Vercel serverless function, supports xlsx and csv with improved detection)
import Busboy from "busboy";
import ExcelJS from "exceljs";
import { Readable } from "stream";
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
          filename: filename || '',
        };
      });
    });
    busboy.on("finish", resolve);
    busboy.on("error", reject);
    req.pipe(busboy);
  });

  await parseForm;

  const mainWorkbook = new ExcelJS.Workbook();
  await mainWorkbook.xlsx.load(buffers.main.buffer);

  const doorKeys = ["door1", "door2", "door3", "door4", "door5", "door6"];
  const doorDataList = [];

  for (const key of doorKeys) {
    const fileData = buffers[key];
    const buffer = fileData.buffer;
    const filename = fileData.filename || '';
    console.log(`Processing ${key}: ${filename}`);

    if (filename.toLowerCase().endsWith(".csv")) {
      const csv = buffer.toString("utf-8");
      const records = parse(csv, { skip_empty_lines: true });
      const names = records.map(([last, first]) => ({ last, first }));
      doorDataList.push(names);
    } else {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      const sheet = wb.worksheets[0];
      const names = [];
      sheet.eachRow((row, rowNum) => {
        if (rowNum >= 4) {
          const last = row.getCell(1).value?.toString().trim();
          const first = row.getCell(2).value?.toString().trim();
          if (last || first) names.push({ last, first });
        }
      });
      doorDataList.push(names);
    }
  }

  for (let i = 0; i < 6; i++) {
    const sheet = mainWorkbook.worksheets[i];
    const doorNames = doorDataList[i];

    // Delete columns A–C from row 4 down
    for (let rowNum = sheet.rowCount; rowNum >= 4; rowNum--) {
      sheet.getRow(rowNum).splice(1, 3);
    }

    // Delete blank rows in Column A & B from row 4
    for (let rowNum = sheet.rowCount; rowNum >= 4; rowNum--) {
      const row = sheet.getRow(rowNum);
      const a = row.getCell(1).value?.toString().trim();
      const b = row.getCell(2).value?.toString().trim();
      if (!a && !b) sheet.spliceRows(rowNum, 1);
    }

    // Paste door names into D & E
    for (let j = 0; j < doorNames.length; j++) {
      const rowNum = j + 4;
      const row = sheet.getRow(rowNum);
      row.getCell(4).value = doorNames[j].last || "";
      row.getCell(5).value = doorNames[j].first || "";
      row.commit();
    }

    // Compare and highlight
    for (let rowNum = 4; rowNum <= sheet.rowCount; rowNum++) {
      const row = sheet.getRow(rowNum);
      const a = (row.getCell(1).value || "").toString().trim();
      const b = (row.getCell(2).value || "").toString().trim();
      const d = (row.getCell(4).value || "").toString().trim();
      const e = (row.getCell(5).value || "").toString().trim();

      const nameOld = `${a} ${b}`.trim();
      const nameNew = `${d} ${e}`.trim();

      if (!nameOld && nameNew) {
        row.getCell(4).fill = row.getCell(5).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF00" },
        };
      }

      if (nameOld && !nameNew) {
        row.getCell(1).fill = row.getCell(2).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0000" },
        };
      }
    }
  }

  const buffer = await mainWorkbook.xlsx.writeBuffer();
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", 'attachment; filename="highlighted-report.xlsx"');
  res.send(Buffer.from(buffer));
}
