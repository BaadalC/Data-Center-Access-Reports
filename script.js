document.getElementById("uploadForm").addEventListener("submit", async function (e) {
  e.preventDefault();
  const statusDiv = document.getElementById("status");
  statusDiv.textContent = "Processing...";

  const mainFile = document.getElementById("mainFile").files[0];
  const doorFiles = Array.from(document.getElementsByClassName("doorFile"));

  if (!mainFile || doorFiles.length !== 6) {
    statusDiv.textContent = "Please upload all required files.";
    return;
  }

  const doorMap = {};
  for (const fileInput of doorFiles) {
    const door = fileInput.dataset.door;
    doorMap[door] = fileInput.files[0];
  }

  try {
    const workbook = XLSX.read(await mainFile.arrayBuffer(), { type: "buffer" });
    const doorNumbers = Object.keys(doorMap);

    for (let i = 0; i < doorNumbers.length; i++) {
      const door = doorNumbers[i];
      const csvFile = doorMap[door];
      const csvText = await csvFile.text();
      const parsed = Papa.parse(csvText.trim(), { header: false }).data;
      const csvNames = parsed
        .slice(1) // skip header
        .filter(row => row[0] || row[1])
        .map(row => [row[0]?.trim().toLowerCase() || "", row[1]?.trim().toLowerCase() || ""]);

      const sheetName = workbook.SheetNames[i];
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const headerRows = data.slice(0, 3);
      let working = data.slice(3)
        .map(r => r.slice(0, 5).concat(["", "", "", "", ""])); // clear F–J
      working.forEach(r => {
        r[0] = ""; r[1] = ""; r[2] = ""; // clear A–C
      });

      const originalNames = data.slice(3).map(r => [
        r[3]?.toString().trim().toLowerCase() || "",
        r[4]?.toString().trim().toLowerCase() || "",
      ]).filter(r => r[0] || r[1]);

      const aligned = alignAndCompare(originalNames, csvNames);
      const finalRows = aligned.map(({ left, right, status }) => {
        const row = [
          left[0] || "",
          left[1] || "",
          "",
          right[0] || "",
          right[1] || "",
          "", "", "", "", "",
        ];

        if (status === "added") {
          row[3] = { v: row[3], s: { fill: { fgColor: { rgb: "FFFF00" } } } };
          row[4] = { v: row[4], s: { fill: { fgColor: { rgb: "FFFF00" } } } };
        } else if (status === "removed") {
          row[0] = { v: row[0], s: { fill: { fgColor: { rgb: "FF0000" } } } };
          row[1] = { v: row[1], s: { fill: { fgColor: { rgb: "FF0000" } } } };
        }

        return row;
      });

      const sheetOutput = XLSX.utils.aoa_to_sheet([...headerRows, ...finalRows]);
      workbook.Sheets[sheetName] = sheetOutput;
    }

    const outputBlob = XLSX.write(workbook, { type: "blob", bookType: "xlsx" });
    const url = URL.createObjectURL(outputBlob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "highlighted_access_report.xlsx";
    link.click();

    statusDiv.textContent = "✅ Report generated successfully!";
  } catch (err) {
    console.error(err);
    statusDiv.textContent = "❌ Something went wrong.";
  }
});

function alignAndCompare(original, incoming) {
  const aligned = [];
  let i = 0, j = 0;

  while (i < original.length || j < incoming.length) {
    const left = i < original.length ? original[i] : ["", ""];
    const right = j < incoming.length ? incoming[j] : ["", ""];

    if (left[0] === right[0] && left[1] === right[1]) {
      aligned.push({ left, right, status: "same" });
      i++; j++;
    } else if (
      left[0] < right[0] || 
      (left[0] === right[0] && left[1] < right[1])
    ) {
      aligned.push({ left, right: ["", ""], status: "removed" });
      i++;
    } else {
      aligned.push({ left: ["", ""], right, status: "added" });
      j++;
    }
  }

  return aligned;
}
