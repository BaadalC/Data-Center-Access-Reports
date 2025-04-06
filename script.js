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

      const clean = str => (str || "").toString().replace(/\*+/g, "").trim();

      const leftNames = workingRows.map(row => [
        clean(row[3]),
        clean(row[4])
      ]).filter(([a, b]) => a || b);

      const aligned = [];
      let iL = 0, iR = 0;
      while (iL < leftNames.length || iR < csvNames.length) {
        const l = iL < leftNames.length ? leftNames[iL] : ["", ""];
        const r = iR < csvNames.length ? csvNames[iR] : ["", ""];

        const lName = `${l[0].toLowerCase()} ${l[1].toLowerCase()}`;
        const rName = `${r[0].toLowerCase()} ${r[1].toLowerCase()}`;

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

      aligned.forEach(obj => {
        const aBlank = !obj.A || obj.A.trim() === "";
        const bBlank = !obj.B || obj.B.trim() === "";
        const dFilled = obj.D && obj.D.trim() !== "";
        const eFilled = obj.E && obj.E.trim() !== "";
        if (aBlank && bBlank && dFilled && eFilled && !obj.highlight) {
          obj.highlight = "added";
        }
      });

      const finalRows = aligned.map(obj => {
        const row = ["", "", "", "", "", "", "", "", "", ""];

        if (obj.highlight === "removed") {
          row[0] = obj.A + "**";
          row[1] = obj.B + "**";
          row[3] = obj.D;
          row[4] = obj.E;
        } else if (obj.highlight === "added") {
          row[0] = obj.A;
          row[1] = obj.B;
          row[3] = obj.D + "*";
          row[4] = obj.E + "*";
        } else {
          row[0] = obj.A;
          row[1] = obj.B;
          row[3] = obj.D;
          row[4] = obj.E;
        }

        return row;
      });

      const finalSheet = XLSX.utils.aoa_to_sheet([...header, ...finalRows]);
      workbook.Sheets[sheetName] = finalSheet;
    }

    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);

    // Naming based on original file, +1 month
    const originalName = mainFile.name;
    const match = originalName.match(/^(\d{4})_(\d{2})/);
    let fileName = "Updated_Access_Report.xlsx";

    if (match) {
      let year = parseInt(match[1]);
      let month = parseInt(match[2]);

      month += 1;
      if (month > 12) {
        month = 1;
        year += 1;
      }

      const mm = String(month).padStart(2, "0");
      fileName = `${year}_${mm}_Data Center Security List.xlsx`;
    }

    link.download = fileName;
    link.click();

    status.textContent = "✅ Report generated!";
  } catch (err) {
    console.error(err);
    status.textContent = "❌ Error occurred.";
  }
});
