document.getElementById("uploadForm").addEventListener("submit", async function (e) {
  e.preventDefault();
  const status = document.getElementById("status");
  status.innerText = "Processing...";

  try {
    const formData = new FormData(this);
    const mainFile = document.getElementById("mainFile").files[0];

    if (!mainFile) {
      throw new Error("Main file is required.");
    }

    const mainWorkbook = XLSX.read(await mainFile.arrayBuffer(), { type: "array" });

    const doorNumbers = ["470", "471", "473", "474", "476", "477"];

    for (const door of doorNumbers) {
      const csvInput = document.querySelector(`.doorFile[data-door="${door}"]`);
      const csvFile = csvInput?.files?.[0];
      if (!csvFile) continue;

      const csvText = await csvFile.text();
      const csvLines = csvText.trim().split("\n").slice(1); // Skip header

      const csvData = Array.from(new Set(
        csvLines.map(row => row.trim()).filter(Boolean)
      ))
        .map(row => row.split(",").map(cell => cell.trim()))
        .filter(row => row.length >= 2 && row[0] && row[1]);

      const sheetName = `Door ${door}`;
      const sheet = mainWorkbook.Sheets[sheetName];
      if (!sheet) continue;

      const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const header = sheetData.slice(0, 3);
      const original = sheetData.slice(3).map(row => {
        const last = row[0]?.toString().trim() || "";
        const first = row[1]?.toString().trim() || "";
        return [last, first];
      }).filter(([a, b]) => a && b);

      const sortedOriginal = [...original].sort();
      const sortedCSV = [...csvData].sort();

      const aligned = [];
      let i = 0, j = 0;

      while (i < sortedOriginal.length || j < sortedCSV.length) {
        const left = i < sortedOriginal.length ? sortedOriginal[i] : ["", ""];
        const right = j < sortedCSV.length ? sortedCSV[j] : ["", ""];

        if (left[0] === right[0] && left[1] === right[1]) {
          aligned.push([left[0], left[1], "", right[0], right[1], "", "", "", "", ""]);
          i++; j++;
        } else if (
          j < sortedCSV.length &&
          (i >= sortedOriginal.length || right[0] < left[0] || (right[0] === left[0] && right[1] < left[1]))
        ) {
          aligned.push(["", "", "", right[0], right[1], "NEW", "", "", "", ""]);
          j++;
        } else {
          aligned.push([left[0], left[1], "", "", "", "REMOVED", "", "", "", ""]);
          i++;
        }
      }

      // Clear F–J from row 4 onward
      aligned.forEach(row => {
        row[5] = row[5] || "";
        row[6] = ""; row[7] = ""; row[8] = ""; row[9] = "";
      });

      const finalData = [...header, ...aligned];
      const updatedSheet = XLSX.utils.aoa_to_sheet(finalData);
      mainWorkbook.Sheets[sheetName] = updatedSheet;
    }

    const updatedBlob = XLSX.write(mainWorkbook, { bookType: "xlsx", type: "blob" });
    const url = URL.createObjectURL(updatedBlob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "Updated_Access_Report.xlsx";
    link.click();
    URL.revokeObjectURL(url);

    status.innerText = "✅ Done!";
  } catch (err) {
    console.error(err);
    status.innerText = "❌ Something went wrong. Please try again.";
  }
});
