document.getElementById("uploadForm").addEventListener("submit", async function (e) {
  e.preventDefault();
  const status = document.getElementById("status");
  status.textContent = "Processing...";

  const mainFile = document.getElementById("mainFile").files[0];
  const doorInputs = document.querySelectorAll(".doorFile");

  if (!mainFile || doorInputs.length === 0) {
    status.textContent = "Please upload the main Excel file and all door CSV files.";
    return;
  }

  try {
    const mainArrayBuffer = await mainFile.arrayBuffer();
    const workbook = XLSX.read(mainArrayBuffer, { type: "array" });

    for (let i = 0; i < doorInputs.length; i++) {
      const input = doorInputs[i];
      const doorNumber = input.getAttribute("data-door");
      const csvFile = input.files[0];
      if (!csvFile) continue;

      const csvText = await csvFile.text();
      const csvParsed = Papa.parse(csvText.trim(), { header: false }).data;

      let csvRaw = csvParsed
        .slice(1)
        .map(row => [row[0]?.trim(), row[1]?.trim(), row[2]?.trim()])
        .filter(([a, b]) => a || b);

      // ✅ Smart deduplication with card number logic
      const uniqueMap = new Map();
      for (const [last, first, card] of csvRaw) {
        const key = `${last.toLowerCase()} ${first.toLowerCase()}`;
        if (!uniqueMap.has(key)) {
          uniqueMap.set(key, [{ last, first, card }]);
        } else {
          uniqueMap.get(key).push({ last, first, card });
        }
      }

      const csvNames = [];
      for (const entries of uniqueMap.values()) {
        if (entries.length === 1) {
          csvNames.push([entries[0].last, entries[0].first]);
        } else {
          const withCard = entries.filter(e => e.card);
          if (withCard.length > 0) {
            for (const e of withCard) {
              csvNames.push([e.last, e.first]);
            }
          } else {
            csvNames.push([entries[0].last, entries[0].first]);
          }
        }
      }

      const sheetName = `Door ${doorNumber}`;
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      const header = data.slice(0, 3);
      const workingRows = data.slice(3);

      const leftNames = workingRows.map(row => [
        (row[3] || "").toString().trim(),
        (row[4] || "").toString().trim()
      ]).filter(([a, b]) => a || b);

      const aligned = [];
      let iL = 0, iR = 0;
      while (iL < leftNames.length || iR < csvNames.length) {
        const l = iL < leftNames.length ? leftNames[iL] : ["", ""];
        const r = iR < csvNames.length ? csvNames[iR] : ["", ""];

        const lName = l[0].toLowerCase() + " " + l[1].toLowerCase();
        const rName = r[0].toLowerCase() + " " + r[1].toLowerCase();

        if (lName === rName) {
          aligned.push({ A: l[0], B: l[1], D: r[0], E: r[1], highlight: null });
          iL++; iR++;
        } else if (lName < rName) {
          aligned.push({ A: l[0], B: l[1], D: "", E: "", highlight: "removed" });
          iL++;
        } else {
          aligned.push({ A: "", B: "", D: r[0], E: r[1], highlight: "added" });
          iR++;
        }
      }

      // ✅ NEW: Re-check for added names where A & B are empty, D & E are filled
      aligned.forEach(obj => {
        const aBlank = !obj.A || obj.A.trim() === "";
        const bBlank = !obj.B || obj.B.trim() === "";
        const dFilled = obj.D && obj.D.trim() !== "";
        const eFilled = obj.E && obj.E.trim() !== "";
        if (aBlank && bBlank && dFilled && eFilled) {
          obj.highlight = "added";
        }
      });

      const red = { fill: { fgColor: { rgb: "FF0000" } } };
      const yellow = { fill: { fgColor: { rgb: "FFFF00" } } };

      const finalRows = aligned.map(obj => {
        const row = [obj.A, obj.B, "", obj.D, obj.E, "", "", "", "", ""];
        if (obj.highlight === "removed") {
          row[0] = { v: obj.A, s: red };
          row[1] = { v: obj.B, s: red };
        } else if (obj.highlight === "added") {
          row[3] = { v: obj.D, s: yellow };
          row[4] = { v: obj.E, s: yellow };
        }
        return row;
      });

      const finalSheet = XLSX.utils.aoa_to_sheet([...header, ...finalRows]);
      workbook.Sheets[sheetName] = finalSheet;
    }

    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Updated_Access_Report.xlsx";
    link.click();

    status.textContent = "✅ File generated!";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Error occurred.";
  }
});
