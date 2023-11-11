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
  if (!excelForm || !fileInput || !divAlert || !btnConvert) {
    throw new Error("One or more required elements not found.");
  }
  var isLoading = false;
  var hideAlert = () => {
    divAlert.style.display = "none";
  };
  var showAlert = (msg, type = "success") => {
    const ALERT_SUCCESS_CLASS = "alert-success";
    const ALERT_ERROR_CLASS = "alert-error";
    let cls = ALERT_SUCCESS_CLASS;
    divAlert.innerHTML = msg;
    divAlert.style.display = "block";
    if (type === "error") {
      divAlert.classList.remove(ALERT_SUCCESS_CLASS);
      cls = "alert-error";
    } else {
      divAlert.classList.remove(ALERT_ERROR_CLASS);
    }
    divAlert.classList.add(cls);
  };
  var toggleSpinner = (el, val) => {
    const LOADING_CLASS = "aria-busy";
    if (isLoading) {
      el.setAttribute(LOADING_CLASS, "true");
    } else {
      el.removeAttribute(LOADING_CLASS);
    }
    el.textContent = val;
  };
  hideAlert();
  toggleSpinner(btnConvert, "Convert");
  excelForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    hideAlert();
    isLoading = true;
    toggleSpinner(btnConvert, "Converting...");
    const selectedFile = fileInput.files?.[0];
    if (!selectedFile) {
      throw new Error("No file selected for conversion.");
    }
    const formData = new FormData();
    formData.append("excelFile", selectedFile);
    try {
      const CONVERT_URL = "/.netlify/functions/convert";
      const res = await fetch(CONVERT_URL, {
        method: "POST",
        body: formData
      });
      if (!res.ok) {
        const err = "Failed to convert. Server returned:";
        console.error(err, res.status, res.statusText);
        showAlert(err, "error");
        throw new Error(err);
      }
      showAlert("Conversion successful. Download will start automatically.");
      isLoading = false;
      toggleSpinner(btnConvert, "Convert");
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
      const msg = "ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and that the layout was not altered.";
      console.error(msg, error);
      showAlert(msg, "error");
      isLoading = false;
      toggleSpinner(btnConvert, "Convert");
    }
  });
})();
