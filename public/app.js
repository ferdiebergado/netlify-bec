"use strict";
(() => {
  // src/app.ts
  var excelForm = document.forms.namedItem(
    "excelForm"
  );
  var fileInput = document.getElementById(
    "excelFile"
  );
  var divAlert = document.getElementById("alert");
  var btnConvert = document.getElementById(
    "convert"
  );
  var hideAlert = () => {
    if (divAlert) {
      divAlert.style.display = "none";
    }
  };
  var showAlert = (msg) => {
    if (divAlert) {
      divAlert.innerHTML = msg;
      divAlert.style.display = "block";
    }
  };
  var toggleSpinner = (el, val) => {
    if (el) {
      if (isLoading) {
        el.setAttribute("aria-busy", "true");
      } else {
        el.removeAttribute("aria-busy");
      }
      el.textContent = val;
    }
  };
  if (!excelForm || !fileInput) {
    throw new Error("Form or file input not found.");
  }
  hideAlert();
  var isLoading = false;
  toggleSpinner(btnConvert, "Convert");
  excelForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    hideAlert();
    isLoading = true;
    toggleSpinner(btnConvert, "Converting...");
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
        showAlert(err);
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
      hideAlert();
      isLoading = false;
      toggleSpinner(btnConvert, "Convert");
    } catch (error) {
      const msg = `ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and the layout was not changed.`;
      console.error(msg, error);
      showAlert(msg);
      isLoading = false;
      toggleSpinner(btnConvert, "Convert");
    }
  });
})();
