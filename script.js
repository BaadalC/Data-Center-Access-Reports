document.getElementById("uploadForm").addEventListener("submit", function (e) {
  e.preventDefault();

  const status = document.getElementById("status");
  status.textContent = "Processing...";

  const mainFile = document.getElementById("mainFile").files[0];
  const doorFiles = document.querySelectorAll(".doorFile");

  if (!mainFile) {
    status.textContent = "Please upload the main working Excel file.";
    return;
  }

  const reader = new FileReader();

  reader.onload = async function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Loop through door files and corresponding workbook sheets
      for (let doorInput of doorFiles) {
        const doorNumber = doorInput.getAttribute("data-door");
        const doorFile = doorInput.files[0];
        if (!doorFile) continue;

        const sheetName = `Door ${doorNumber}`;
        const worksheet = workbook.Sheets[sheetName];

        if (!worksheet) {
          console.warn(`Sheet "${sheetName}" not found in Excel.`);
          continue;
        }

        // Parse CSV
        const csvText = await doorFile.text();
        const csvData = XLSX.read(csvText, { type: "string" });
        const csvSheet = csvData.Sheets[csvData.SheetNames[0]];
        const csvJson = XLSX.utils.sheet_to_json(csvSheet, { header: 1 });

        const csvNames = csvJson.slice(1).map((row) => ({
          last: (row[0] || "").toString().trim(),
          first: (row[1] || "").toString().trim(),
        }));

        // Clean up existing content from columns A–C (leave rows 1–3 alone)
        const sheetJson = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const cleanedRows = sheetJson.map((row, i) =>
          i <= 2 ? row : row.slice(3)
        );

        // Remove empty rows after cleaning
        const filteredRows = cleanedRows.filter((row, i) =>
          i <= 2 ? true : row[0] || row[1]
        );

        // Add CSV data into columns D and E, starting at row 4
        for (let i = 0; i < csvNames.length; i++) {
          const targetRow = filteredRows[i + 3] || [];
          targetRow[3] = csvNames[i].last;
          targetRow[4] = csvNames[i].first;
          filteredRows[i + 3] = targetRow;
        }

        // Align names and insert blanks where needed
        const alignedRows = alignRows(filteredRows);

        // Clear F–J from row 4 onward
        for (let i = 3; i < alignedRows.length; i++) {
          alignedRows[i][5] = "";
          alignedRows[i][6] = "";
          alignedRows[i][7] = "";
          alignedRows[i][8] = "";
          alignedRows[i][9] = "";
        }

        // Write back to worksheet
        const updatedSheet = XLSX.utils.aoa_to_sheet(alignedRows);
        workbook.Sheets[sheetName] = updatedSheet;
      }

      // Generate downloadable file
      const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/octet-stream" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "Updated_Access_Report.xlsx";
      link.click();

      status.textContent = "✅ Done! File downloaded.";
    } catch (err) {
      console.error(err);
      status.textContent = "❌ Something went wrong.";
    }
  };

  reader.readAsArrayBuffer(mainFile);
});

// Alignment helper (you can customize this further)
function alignRows(rows) {
  const result = [rows[0], rows[1], rows[2]];

  let i = 3;
  while (i < rows.length) {
    const row = rows[i];
    const left = `${row[0] || ""} ${row[1] || ""}`.trim().toLowerCase();
    const right = `${row[3] || ""} ${row[4] || ""}`.trim().toLowerCase();

    if (left === right) {
      result.push(row);
      i++;
    } else {
      const nextRight = (rows[i + 1] || []);
      const nextRightName = `${nextRight[3] || ""} ${nextRight[4] || ""}`.trim().toLowerCase();
      if (left === nextRightName) {
        // New name added in right, insert blank in A/B
        result.push(["", "", ...(row.slice(2))]);
      } else {
        // Name removed, insert blank in D/E
        result.push([...(row.slice(0, 2)), "", "", ...(row.slice(5))]);
      }
      i++;
    }
  }

  return result;
}
