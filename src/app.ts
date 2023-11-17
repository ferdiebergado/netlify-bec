import { CONVERT_URL } from './server/constants';
import { createTimestamp } from './server/utils';

const excelForm = document.getElementById(
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

function hideAlert() {
  divAlert!.style.display = 'none';
}

function showAlert(msg: string, type: string = 'success') {
  const ALERT_SUCCESS_CLASS = 'alert-success';
  const ALERT_ERROR_CLASS = 'alert-error';

  let cls = ALERT_SUCCESS_CLASS;

  divAlert!.innerHTML = msg;
  divAlert!.style.display = 'block';

  if (type === 'error') {
    divAlert!.classList.remove(ALERT_SUCCESS_CLASS);
    cls = 'alert-error';
  } else {
    divAlert!.classList.remove(ALERT_ERROR_CLASS);
  }

  divAlert!.classList.add(cls);
}

function toggleSpinner(el: HTMLElement, val: string) {
  const LOADING_CLASS = 'aria-busy';

  if (isLoading) {
    el.setAttribute(LOADING_CLASS, 'true');
  } else {
    el.removeAttribute(LOADING_CLASS);
  }

  el.textContent = val;
}

async function handleSubmit(event: SubmitEvent) {
  event.preventDefault();

  hideAlert();
  isLoading = true;
  toggleSpinner(btnConvert!, 'Converting...');

  const selectedFile = fileInput!.files?.[0];

  if (!selectedFile) {
    throw new Error('No file selected for conversion.');
  }

  const formData = new FormData();
  formData.append('excelFile', selectedFile);

  try {
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
    toggleSpinner(btnConvert!, 'Convert');

    // Get the blob data from the response
    const blob = await res.blob();

    // Get the filename from the Content-Disposition header
    const contentDisposition = res.headers.get('Content-Disposition');
    const filenameMatch =
      contentDisposition && contentDisposition.match(/filename="(.+?)"/);
    const filename = filenameMatch
      ? filenameMatch[1]
      : `em-${createTimestamp()}.xlsx`;

    // Create a Blob URL
    const blobUrl = URL.createObjectURL(blob);

    // Create an anchor element to trigger the download
    const a = document.createElement('a');
    a.href = blobUrl;
    a.download = filename;

    // Append the anchor to the document and trigger the download
    document.body.appendChild(a);
    a.click();

    // Clean up: remove the anchor and revoke the Blob URL
    document.body.removeChild(a);
    URL.revokeObjectURL(blobUrl);
  } catch (error) {
    const msg =
      'ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and that the layout was not altered.';
    console.error(msg, error);
    // Handle error as needed
    showAlert(msg, 'error');
    isLoading = false;
    toggleSpinner(btnConvert!, 'Convert');
  }
}

hideAlert();

toggleSpinner(btnConvert, 'Convert');

excelForm.addEventListener('submit', handleSubmit);
