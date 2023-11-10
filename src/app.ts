const btnConvert = document.getElementById("convert");
const fileInput = document.getElementById("excelFile");

btnConvert?.addEventListener("click", async (event) => {
  event.preventDefault();
  const formData = new FormData();
  formData.append("excelFile", <HTMLInputElement>fileInput?.files[0]);
  const res = await fetch("/convert", {
    method: "POST",
    body: formData,
  });
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  window.location.href = url;
});
