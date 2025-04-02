import { read, utils, writeFile } from "xlsx";

function normalizeName(name) {
  return name ? name.toString().trim().toLowerCase() : "";
}

function alignAndCompare(original, csvData) {
  const originalClean = original
    .filter(row => normalizeName(row[0]) || normalizeName(row[1]))
    .map(row => [normalizeName(row[0]), normalizeName(row[1])]);

  const csvClean = csvData
    .filter(row => normalizeName(row[0]) || normalizeName(row[1]))
    .map(row => [normalizeName(row[0]), normalizeName(row[1])]);

  let i = 0, j = 0;
  const result = [];

  while (i < originalClean.length || j < csvClean.length) {
    const left = i < originalClean.length ? originalClean[i] : ["", ""];
    const right = j < csvClean.length ? csvClean[j] : ["", ""];

    if (left[0] === right[0] && left[1] === right[1]) {
      result.push({ left, right, status: "same" });
      i++; j++;
    } else if (
      left[0] < right[0] || 
      (left[0] === right[0] && left[1] < right[1])
    ) {
      result.push({ left, right: ["", ""], status: "removed" });
      i++;
    } else {
      result.push({ left: ["", ""], right, status: "added" });
      j++;
    }
  }

  return result;
}

export async function processExcel(mainFile, doorCSVs) {
  const mainWorkbook = read(await mainFile.arrayBuffer(), { type: "buffer" });

  const csvBuffers = await Promise.all(
    doorCSVs.map(file => file.arrayBuffer())
  );
  const csvWorksheets = csvBuffers.map(buffer =>
    utils.sheet_to_json(read(buffer, { type: "buffer", raw: false }).Sheets.Sheet1, {
      header: 1,
      blankrows: false,
    })
  );

  doorCSVs.forEach((file, idx) => {
    const sheetName = mainWorkbook.SheetNames[idx];
    const sheet = mainWorkbook.Sheets[sheetName];
    const sheetData = utils.sheet_to_json(sheet, { header: 1, blankrows: false });

    const headers = sheetData.slice(0, 3);
    let workingData = sheetData.slice(3);

    // Remove extra columns (F–J) and highlights
    workingData = workingData.map(row => {
      for (let i = 5; i < 10; i++) row[i] = "";
      return row.slice(0, 5);
    });

    // Clear columns A–C from row 4
    workingData = workingData.map(row => ["", "", "", row[3] || "", row[4] || ""]);

    // Align names
    const aligned = alignAndCompare(workingData.map(r => [r[0], r[1]]), csvWorksheets[idx]);

    const updated = aligned.map(item => {
      const row = [
        item.left[0] || "",
        item.left[1] || "",
        "", // Column C left blank
        item.right[0] || "",
        item.right[1] || "",
      ];

      if (item.status === "added") {
        row.push({ s: { fill: { fgColor: { rgb: "FFFF00" } } } }); // yellow
      } else if (item.status === "removed") {
        row[0] = { v: row[0], s: { fill: { fgColor: { rgb: "FF0000" } } } };
        row[1] = { v: row[1], s: { fill: { fgColor: { rgb: "FF0000" } } } };
      }

      return row;
    });

    const finalData = [...headers, ...updated];
    const newSheet = utils.aoa_to_sheet(finalData);
    mainWorkbook.Sheets[sheetName] = newSheet;
  });

  const outputBlob = writeFile(mainWorkbook, "highlighted_result.xlsx", {
    bookType: "xlsx",
    type: "binary",
  });

  return outputBlob;
}
