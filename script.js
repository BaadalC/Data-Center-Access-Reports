document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("uploadForm");

  form.addEventListener("submit", async (e) => {
    e.preventDefault();

    const status = document.getElementById("status");
    status.innerText = "Processing...";

    const formData = new FormData();

    try {
      formData.append("mainFile", document.getElementById("mainFile").files[0]);
      formData.append("door470", document.getElementById("door470").files[0]);
      formData.append("door471", document.getElementById("door471").files[0]);
      formData.append("door473", document.getElementById("door473").files[0]);
      formData.append("door474", document.getElementById("door474").files[0]);
      formData.append("door476", document.getElementById("door476").files[0]);
      formData.append("door477", document.getElementById("door477").files[0]);
    } catch (error) {
      status.innerText = "Please select all required files.";
      return;
    }

    try {
      const response = await fetch("/api/compare", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Server error");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "highlighted.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      status.innerText = "✅ Download ready!";
    } catch (error) {
      status.innerText = "❌ Something went wrong. Please try again.";
      console.error(error);
    }
  });
});
