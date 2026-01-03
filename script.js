pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const fileInput = document.getElementById("pdfFile");
const convertBtn = document.getElementById("convertBtn");
const statusText = document.getElementById("status");

let pdfFile = null;

fileInput.addEventListener("change", () => {
  pdfFile = fileInput.files[0];
  convertBtn.disabled = !pdfFile;
});

convertBtn.addEventListener("click", async () => {
  if (!pdfFile) return;

  statusText.textContent = "Reading PDF...";

  const reader = new FileReader();
  reader.readAsArrayBuffer(pdfFile);

  reader.onload = async () => {
    try {
      const typedarray = new Uint8Array(reader.result);
      const pdf = await pdfjsLib.getDocument({ data: typedarray }).promise;

      let rows = [];

      for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const content = await page.getTextContent();

        let line = [];
        let lastY = null;

        content.items.forEach(item => {
          const y = item.transform[5];

          if (lastY === null || Math.abs(y - lastY) < 4) {
            line.push(item.str);
          } else {
            rows.push(line);
            line = [item.str];
          }

          lastY = y;
        });

        if (line.length) rows.push(line);
      }

      if (!rows.length) {
        statusText.textContent = "No readable text found in PDF.";
        return;
      }

      statusText.textContent = "Generating Excel...";

      const workbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(rows);

      XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Report");
      XLSX.writeFile(workbook, "sales-report.xlsx");

      statusText.textContent = "Done! Excel downloaded.";

    } catch (err) {
      console.error(err);
      statusText.textContent = "Error parsing PDF.";
    }
  };
});
