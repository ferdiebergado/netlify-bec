const excelForm = document.forms.namedItem(
  "excelForm"
) as HTMLFormElement | null;
const fileInput = document.getElementById(
  "excelFile"
) as HTMLInputElement | null;

if (!excelForm || !fileInput) {
  throw new Error("Form or file input not found.");
  // Handle case where form or file input is not found
}

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
      body: formData,
    });

    if (!res.ok) {
      const err = "Failed to convert. Server returned:";
      console.error(err, res.status, res.statusText);
      throw new Error(err);
    }

    // Get the blob data from the response
    const blob = await res.blob();

    // Get the filename from the Content-Disposition header
    const contentDisposition = res.headers.get("Content-Disposition");
    const filenameMatch =
      contentDisposition && contentDisposition.match(/filename="(.+?)"/);
    const filename = filenameMatch
      ? filenameMatch[1]
      : `em-${new Date().getTime()}.xlsx`;

    // Create a download link
    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = filename;

    // Append the link to the body and trigger the click event
    document.body.appendChild(downloadLink);
    downloadLink.click();

    // // Clean up: remove the link from the body
    document.body.removeChild(downloadLink);
  } catch (error) {
    console.error("An error occurred during conversion:", error);
    // Handle error as needed
  }
});
