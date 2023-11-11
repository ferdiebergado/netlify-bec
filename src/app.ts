const excelForm = document.forms.namedItem(
  'excelForm',
) as HTMLFormElement | null;
const fileInput = document.getElementById(
  'excelFile',
) as HTMLInputElement | null;
const divAlert = document.getElementById('alert') as HTMLDivElement | null;
const btnConvert = document.getElementById(
  'convert',
) as HTMLButtonElement | null;

if (!excelForm || !fileInput || !divAlert || !btnConvert) {
  throw new Error('One or more required elements not found.');
}

let isLoading = false;

const hideAlert = () => {
  divAlert.style.display = 'none';
};

const showAlert = (msg: string, type: string = 'success') => {
  const ALERT_SUCCESS_CLASS = 'alert-success';
  const ALERT_ERROR_CLASS = 'alert-error';

  let cls = ALERT_SUCCESS_CLASS;

  divAlert.innerHTML = msg;
  divAlert.style.display = 'block';

  if (type === 'error') {
    divAlert.classList.remove(ALERT_SUCCESS_CLASS);
    cls = 'alert-error';
  } else {
    divAlert.classList.remove(ALERT_ERROR_CLASS);
  }

  divAlert.classList.add(cls);
};

const toggleSpinner = (el: HTMLElement, val: string) => {
  const LOADING_CLASS = 'aria-busy';

  if (isLoading) {
    el.setAttribute(LOADING_CLASS, 'true');
  } else {
    el.removeAttribute(LOADING_CLASS);
  }

  el.textContent = val;
};

hideAlert();

toggleSpinner(btnConvert, 'Convert');

excelForm.addEventListener('submit', async event => {
  event.preventDefault();

  hideAlert();
  isLoading = true;
  toggleSpinner(btnConvert, 'Converting...');

  const selectedFile = fileInput.files?.[0];

  if (!selectedFile) {
    throw new Error('No file selected for conversion.');
  }

  const formData = new FormData();
  formData.append('excelFile', selectedFile);

  try {
    const CONVERT_URL = '/.netlify/functions/convert';
    const res = await fetch(CONVERT_URL, {
      method: 'POST',
      body: formData,
    });

    if (!res.ok) {
      const err = 'Failed to convert. Server returned:';
      console.error(err, res.status, res.statusText);

      showAlert(err, 'error');

      throw new Error(err);
    }

    showAlert('Conversion successful. Download will start automatically.');
    isLoading = false;
    toggleSpinner(btnConvert, 'Convert');

    // Get the blob data from the response
    const blob = await res.blob();

    // Get the filename from the Content-Disposition header
    const contentDisposition = res.headers.get('Content-Disposition');
    const filenameMatch =
      contentDisposition && contentDisposition.match(/filename="(.+?)"/);
    const filename = filenameMatch
      ? filenameMatch[1]
      : `em-${new Date().getTime()}.xlsx`;

    // Create a download link
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = filename;

    // Append the link to the body and trigger the click event
    document.body.appendChild(downloadLink);
    downloadLink.click();

    // // Clean up: remove the link from the body
    document.body.removeChild(downloadLink);
  } catch (error) {
    const msg =
      'ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and that the layout was not altered.';
    console.error(msg, error);
    // Handle error as needed
    showAlert(msg, 'error');
    isLoading = false;
    toggleSpinner(btnConvert, 'Convert');
  }
});
