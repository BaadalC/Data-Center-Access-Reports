import formidable from 'formidable';
import fs from 'fs';
import { Readable } from 'stream';
import { Workbook } from 'exceljs';
import csv from 'csv-parser';

// Disable default body parser (important for file upload)
export const config = {
  api: {
    bodyParser: false,
  },
};

function parseCSV(buffer) {
  return new Promise((resolve, reject) => {
    const results = [];
    const stream = Readable.from(buffer.toString());

    stream
      .pipe(csv())
      .on('data', (data) => results.push(data))
      .on('end', () => resolve(results))
      .on('error', reject);
  });
}

async function parseFormAsync(req) {
  const form = formidable({ multiples: true });

  return new Promise((resolve, reject) => {
    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      else resolve({ fields, files });
    });
  });
}

export default async function handler(req, res) {
  try {
    if (req.method !== 'POST') {
      return res.status(405).send('Method Not Allowed');
    }

    const { files } = await parseFormAsync(req);

    const mainFile = files.mainFile[0];
    const referenceFiles = [
      files.door470?.[0],
      files.door471?.[0],
      files.door473?.[0],
      files.door474?.[0],
      files.door476?.[0],
      files.door477?.[0]
    ];

    // Load main Excel file
    const workbook = new Workbook();
    await workbook.xlsx.readFile(mainFile.filepath);

    for (let i = 0; i < referenceFiles.length; i++) {
      const refFile = referenceFiles[i];
      if (!refFile) continue;

      const sheetName = `Door ${470 + i}`;
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) continue;

      // Step 6–7: Delete columns A–C starting from row 4
      worksheet.spliceColumns(1, 3);

      // Step 8: Remove empty rows based on column A & B
      let rowIndex = 4;
      while (rowIndex <= worksheet.rowCount) {
        const row = worksheet.getRow(rowIndex);
        const colA = row.getCell(1).value;
        const colB = row.getCell(2).value;

        if (!colA && !colB) {
          worksheet.spliceRows(rowIndex, 1);
        } else {
          rowIndex++;
        }
      }

      // Step 9–10: Paste Door CSV into D&E
      const buffer = fs.readFileSync(refFile.filepath);
      const csvData = await parseCSV(buffer);

      for (let j = 0; j < csvData.length; j++) {
        const nameRow = csvData[j];
        const excelRow = worksheet.getRow(4 + j);
        excelRow.getCell(4).value = nameRow['Last Name'] || '';
        excelRow.getCell(5).value = nameRow['First Name'] || '';
        excelRow.commit();
      }

      // Step 11+: Align rows (skipped for now) + highlight new names
      for (let j = 4; j <= worksheet.rowCount; j++) {
        const cellOld = worksheet.getRow(j).getCell(1).value;
        const cellNew = worksheet.getRow(j).getCell(4).value;
        if (cellNew && !cellOld) {
          worksheet.getRow(j).getCell(4).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF00' } // Yellow
          };
          worksheet.getRow(j).getCell(5).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFF00' }
          };
        }
      }
    }

    // Write final file to buffer
    const outBuffer = await workbook.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=highlighted.xlsx');
    res.send(outBuffer);

  } catch (err) {
    console.error('❌ Server Error:', err);
    res.status(500).send('Internal Server Error: ' + err.message);
  }
}
