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
  const mainFileBuffer = buffers["mainFile"];
  const isMainCsv = mainFileBuffer.toString("utf8").startsWith("First Name") || mainFileBuffer.toString("utf8").startsWith('"First Name"');
  await (isMainCsv
    ? mainWorkbook.csv.read(Readable.from(mainFileBuffer))
    : mainWorkbook.xlsx.load(mainFileBuffer));

  const newWorkbook = new ExcelJS.Workbook();

  for (const field in buffers) {
    if (field === "mainFile") continue;

    const doorNumber = field.replace("door", "");
    const referenceBuffer = buffers[field];

    const referenceWorkbook = new ExcelJS.Workbook();
    await referenceWorkbook.csv.read(Readable.from(referenceBuffer));
    const referenceSheet = referenceWorkbook.worksheets[0];

    const tabName = `Door ${doorNumber}`;
    const mainSheet = mainWorkbook.getWorksheet(tabName);
    if (!mainSheet) continue;

    const newSheet = newWorkbook.addWorksheet(tabName);

    // Copy headers (first 3 rows)
    for (let i = 1; i <= 3; i++) {
      const row = mainSheet.getRow(i);
      newSheet.addRow(row.values);
    }

    const oldNames = new Set();
    mainSheet.eachRow((row, rowNumber) => {
      if (rowNumber > 3 && row.getCell(1).value && row.getCell(2).value) {
        oldNames.add(`${row.getCell(1).value} ${row.getCell(2).value}`);
      }
    });

    const newNames = new Set();
    referenceSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      if (row.getCell(1).value && row.getCell(2).value) {
        newNames.add(`${row.getCell(1).value} ${row.getCell(2).value}`);
      }
    });

    // Highlight additions
    referenceSheet.eachRow((row, rowNumber) => {
      const newRow = [];
      row.eachCell(cell => newRow.push(cell.value));
      const fullName = `${row.getCell(1).value} ${row.getCell(2).value}`;
      const excelRow = newSheet.addRow(newRow);
      if (!oldNames.has(fullName)) {
        excelRow.eachCell(cell => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF00" },
          };
        });
      }
    });

    // Highlight removals
    mainSheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 3) return;
      const fullName = `${row.getCell(1).value} ${row.getCell(2).value}`;
      if (!newNames.has(fullName)) {
        const values = [];
        row.eachCell(cell => values.push(cell.value));
        const removedRow = newSheet.addRow(values);
        removedRow.eachCell(cell => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000" },
          };
        });
      }
    });
  }

  const buffer = await newWorkbook.xlsx.writeBuffer();
  res.setHeader("Content-Disposition", "attachment; filename=highlighted.xlsx");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
}
