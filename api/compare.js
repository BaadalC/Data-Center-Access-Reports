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

  try {
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

      busboy.on("finish", () => resolve());
      busboy.on("error", reject);
    });

    req.pipe(busboy);
    await parseForm();

    // Parse the main Excel file
    const mainWorkbook = new ExcelJS.Workbook();
    await mainWorkbook.xlsx.load(buffers.mainFile);

    // Load reference files (CSV)
    const referenceFiles = {};
    for (const key of Object.keys(buffers)) {
      if (key.startsWith("door")) {
        referenceFiles[key] = buffers[key].toString("utf8").split("\n").map(line => line.trim().split(","));
      }
    }

    // We'll process only Door 470 for now
    const sheet = mainWorkbook.getWorksheet("Door 470");
    if (!sheet) throw new Error("Sheet 'Door 470' not found in main Excel file");

    // Remove columns A to C (from row 4 onward)
    for (let i = sheet.actualRowCount; i >= 4; i--) {
      sheet.getRow(i).splice(1, 3); // A to C is 1 to 3
    }

    // Clean up empty rows in col A & B
    for (let i = sheet.actualRowCount; i >= 4; i--) {
      const row = sheet.getRow(i);
      const valA = row.getCell(1).value;
      const valB = row.getCell(2).value;
      if (!valA && !valB) {
        sheet.spliceRows(i, 1);
      }
    }

    // Paste reference CSV to col D & E (starting row 4)
    const ref = referenceFiles["door470"];
    if (!ref) throw new Error("CSV for Door 470 not found");

    for (let i = 0; i < ref.length; i++) {
      const row = sheet.getRow(i + 4);
      if (!row) continue;
      row.getCell(4).value = ref[i][0]; // Last name
      row.getCell(5).value = ref[i][1]; // First name
    }

    // TODO: highlight new names & align missing (will do this next)
    
    // Send back the Excel
    const buffer = await mainWorkbook.xlsx.writeBuffer();
    res.setHeader("Content-Disposition", "attachment; filename=highlighted.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);

  } catch (err) {
    console.error("💥 Error in compare API:", err);
    res.status(500).send("Internal Server Error: " + err.message);
  }
}
