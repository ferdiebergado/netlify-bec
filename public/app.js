"use strict";
(() => {
  // src/app.ts
  var excelForm = document.forms.namedItem(
    "excelForm"
  );
  var fileInput = document.getElementById(
    "excelFile"
  );
  if (excelForm && fileInput) {
    excelForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = new FormData();
      const selectedFile = fileInput.files?.[0];
      if (!selectedFile) {
        throw new Error("No file selected for conversion.");
      }
      formData.append("excelFile", selectedFile);
      try {
        const res = await fetch("/.netlify/functions/convert", {
          method: "POST",
          body: formData
        });
        if (!res.ok) {
          const err = "Failed to convert. Server returned:";
          console.error(err, res.status, res.statusText);
          throw new Error(err);
        }
        const blob = await res.blob();
        const contentDisposition = res.headers.get("Content-Disposition");
        const filenameMatch = contentDisposition && contentDisposition.match(/filename="(.+?)"/);
        const filename = filenameMatch ? filenameMatch[1] : `em-${(/* @__PURE__ */ new Date()).getTime()}.xlsx`;
        const downloadLink = document.createElement("a");
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = filename;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
      } catch (error) {
        console.error("An error occurred during conversion:", error);
      }
    });
  } else {
    console.error("Form or file input not found.");
  }
})();
