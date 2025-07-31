let selectedFile;

document.getElementById("upload").addEventListener("change", (event) => {
  selectedFile = event.target.files[0];
});

function convertToCSV() {
  if (!selectedFile) {
    alert("Please upload an Excel file first.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const csv = XLSX.utils.sheet_to_csv(sheet);

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const link = document.getElementById("downloadLink");
    link.href = url;
    link.style.display = "inline-block";
  };

  reader.readAsArrayBuffer(selectedFile);
}
