document.getElementById("uploadForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const statusDiv = document.getElementById("status");
  statusDiv.innerHTML = "Processing...";

  try {
    const mainFile = document.getElementById("mainFile").files[0];
    if (!mainFile) throw new Error("Main Excel file not uploaded.");

    const doorFiles = {};
    document.querySelectorAll(".doorFile").forEach((input) => {
      const door = input.dataset.door;
      if (input.files.length > 0) {
        doorFiles[door] = input.files[0];
      }
    });

    const mainWorkbook = await readExcelFile(mainFile);
    for (const [door, file] of Object.entries(doorFiles)) {
      const csvData = await readCSVFile(file);
      processTab(mainWorkbook, door, csvData);
    }

    const outFile = XLSX.write(mainWorkbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([outFile], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "Updated_Access_Report.xlsx";
    a.click();

    statusDiv.innerHTML = "✅ Done!";
  } catch (err) {
    console.error(err);
    statusDiv.innerHTML = "❌ Something went wrong.";
  }
});

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      resolve(workbook);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function readCSVFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      complete: (results) => {
        resolve(results.data.filter(row => row[0] && row[1]));
      },
      error: reject
    });
  });
}

function processTab(workbook, door, csvData) {
  const sheetName = workbook.SheetNames.find(name => name.includes(door));
  if (!sheetName) return;

  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Step 1: Clear columns A–C from row 4 downward and shift left
  data.splice(3, data.length - 3); // Remove old A–C content
  for (let i = 3; i < data.length; i++) {
    data[i].splice(0, 3);
  }

  // Step 2: Remove empty rows
  const filtered = data.slice(0, 3).concat(
    data.slice(3).filter(row => row[0] && row[1])
  );

  // Step 3: Clean any leftover in columns F–J from row 4 down
  for (let i = 3; i < filtered.length; i++) {
    for (let col = 5; col <= 9; col++) {
      filtered[i][col] = "";
    }
  }

  // Step 4: Paste new CSV data into col D/E from row 4
  const leftSide = filtered.slice(3).map(r => [r[0], r[1]]);
  const rightSideRaw = csvData.map(r => [r[0].trim(), r[1].trim()]);

  // Remove duplicates in right side
  const uniqueRightSide = [];
  const seen = new Set();
  for (const [ln, fn] of rightSideRaw) {
    const key = `${ln},${fn}`;
    if (!seen.has(key)) {
      seen.add(key);
      uniqueRightSide.push([ln, fn]);
    }
  }

  // Step 5: Align both sides
  const aligned = [];
  let i = 0, j = 0;
  while (i < leftSide.length || j < uniqueRightSide.length) {
    const left = leftSide[i] || ["", ""];
    const right = uniqueRightSide[j] || ["", ""];

    if (left[0] === right[0] && left[1] === right[1]) {
      aligned.push([left[0], left[1], "", right[0], right[1]]);
      i++; j++;
    } else {
      const lStr = `${left[0]} ${left[1]}`.toLowerCase();
      const rStr = `${right[0]} ${right[1]}`.toLowerCase();

      if (!left[0] && right[0]) {
        aligned.push(["", "", "", right[0], right[1]]);
        j++;
      } else if (!right[0] && left[0]) {
        aligned.push([left[0], left[1], "", "", ""]);
        i++;
      } else if (lStr < rStr) {
        aligned.push([left[0], left[1], "", "", ""]);
        i++;
      } else {
        aligned.push(["", "", "", right[0], right[1]]);
        j++;
      }
    }
  }

  // Step 6: Highlight mismatches
  const ws = XLSX.utils.aoa_to_sheet(filtered.slice(0, 3).concat(aligned));
  const range = XLSX.utils.decode_range(ws["!ref"]);

  for (let row = 4; row <= range.e.r + 1; row++) {
    const leftL = ws[`A${row}`]?.v;
    const leftF = ws[`B${row}`]?.v;
    const rightL = ws[`D${row}`]?.v;
    const rightF = ws[`E${row}`]?.v;

    if (!leftL && rightL) {
      ws[`D${row}`].s = ws[`E${row}`].s = {
        fill: { fgColor: { rgb: "FFFF00" } },
      };
    } else if (leftL && !rightL) {
      ws[`A${row}`].s = ws[`B${row}`].s = {
        fill: { fgColor: { rgb: "FFFF00" } },
      };
    }
  }

  ws["!cols"] = [];
  workbook.Sheets[sheetName] = ws;
}
