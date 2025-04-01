// /api/compare.js (Vercel serverless function)
import Busboy from "busboy";
import ExcelJS from "exceljs";
import { Readable } from "stream";

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

  await parseForm;

  const mainWorkbook = new ExcelJS.Workbook();
  await mainWorkbook.xlsx.load(buffers.main);

  const doorBuffers = ["door1", "door2", "door3", "door4", "door5", "door6"].map(key => buffers[key]);
  const doorWorkbooks = [];
  for (const buffer of doorBuffers) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    doorWorkbooks.push(wb);
  }

  for (let i = 0; i < 6; i++) {
    const sheet = mainWorkbook.worksheets[i];
    const doorSheet = doorWorkbooks[i].worksheets[0];

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

    // Read names from door file (columns A & B)
    const doorNames = [];
    doorSheet.eachRow((row, rowNum) => {
      if (rowNum >= 4) {
        const last = row.getCell(1).value?.toString().trim();
        const first = row.getCell(2).value?.toString().trim();
        if (last || first) doorNames.push({ last, first });
      }
    });

    // Write names into columns D & E of main sheet
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
        // New name → highlight yellow
        row.getCell(4).fill = row.getCell(5).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFF00" },
        };
      }

      if (nameOld && !nameNew) {
        // Removed name → highlight red in A & B
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
