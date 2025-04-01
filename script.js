document.getElementById('uploadForm').addEventListener('submit', async function (e) {
  e.preventDefault();

  const formData = new FormData();
  const files = {
    main: document.getElementById("mainFile").files[0],
    door1: document.getElementById("door1").files[0],
    door2: document.getElementById("door2").files[0],
    door3: document.getElementById("door3").files[0],
    door4: document.getElementById("door4").files[0],
    door5: document.getElementById("door5").files[0],
    door6: document.getElementById("door6").files[0],
  };

  for (const [key, file] of Object.entries(files)) {
    if (!file) {
      alert(`Missing file: ${key}`);
      return;
    }
    formData.append(key, file);
  }

  document.getElementById("status").innerText = "Processing...";

  try {
    const response = await fetch("/api/compare", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) throw new Error("Server error");

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "highlighted-report.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();

    document.getElementById("status").innerText = "Done! Download started.";
  } catch (error) {
    console.error(error);
    document.getElementById("status").innerText = "Something went wrong.";
  }
});
