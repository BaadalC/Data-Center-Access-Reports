
import Busboy from "busboy";
import ExcelJS from "exceljs";
import csvParser from "csv-parser";
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

  try {
    await parseForm();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffers["mainFile"]);

    for (let door = 470; door <= 477; door++) {
      const tabName = "Door " + door;
      const sheet = workbook.getWorksheet(tabName);
      const csvBuffer = buffers[`door${door}`];

      if (!sheet || !csvBuffer) continue;

      // Step 1: Delete columns A-C (1-3) from row 4 down
      const lastRow = sheet.lastRow.number;
      for (let i = 4; i <= lastRow; i++) {
        sheet.getRow(i).splice(1, 3);
      }

      // Step 2: Delete blank rows in A & B
      for (let i = lastRow; i >= 4; i--) {
        const row = sheet.getRow(i);
        if (!row.getCell(1).value && !row.getCell(2).value) {
          sheet.spliceRows(i, 1);
        }
      }

      // Step 3: Parse reference CSV and store names
      const newEntries = [];
      await new Promise((resolve, reject) => {
        Readable.from(csvBuffer)
          .pipe(csvParser())
          .on("data", (row) => {
            newEntries.push({
              lastName: row[Object.keys(row)[0]]?.trim(),
              firstName: row[Object.keys(row)[1]]?.trim(),
            });
          })
          .on("end", resolve)
          .on("error", reject);
      });

      // Sort entries
      newEntries.sort((a, b) => {
        const aName = `${a.lastName} ${a.firstName}`;
        const bName = `${b.lastName} ${b.firstName}`;
        return aName.localeCompare(bName);
      });

      // Insert into D & E
      for (let i = 0; i < newEntries.length; i++) {
        const rowIndex = i + 4;
        const row = sheet.getRow(rowIndex);
        row.getCell(4).value = newEntries[i].lastName || "";
        row.getCell(5).value = newEntries[i].firstName || "";
        row.commit();
      }

      // Highlight logic
      const maxLen = Math.max(sheet.rowCount - 3, newEntries.length);
      for (let i = 0; i < maxLen; i++) {
        const rowIndex = i + 4;
        const row = sheet.getRow(rowIndex);
        const oldName = `${row.getCell(1).value || ""} ${row.getCell(2).value || ""}`.trim();
        const newName = `${row.getCell(4).value || ""} ${row.getCell(5).value || ""}`.trim();

        if (oldName && !newName) {
          // Access removed (RED)
          row.getCell(1).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" },
          };
          row.getCell(2).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF0000" },
          };
        } else if (!oldName && newName) {
          // Access added (YELLOW)
          row.getCell(4).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" },
          };
          row.getCell(5).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFFF00" },
          };
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Disposition", "attachment; filename=highlighted.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.status(200).send(buffer);
  } catch (err) {
    console.error("Error:", err);
    res.status(500).send("Internal Server Error: " + err.message);
  }
}
