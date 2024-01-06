import convert from './converter';
import { createTimestamp } from './utils';

// Constants
const ALERT_SUCCESS_CLASS = 'alert-success';
const ALERT_ERROR_CLASS = 'alert-error';

// Elements
const excelForm = document.forms.namedItem('excelForm');
const fileInput = document.getElementById('excelFile') as HTMLInputElement;
const divAlert = document.getElementById('alert') as HTMLDivElement;
const btnConvert = document.getElementById('convert') as HTMLButtonElement;

let isLoading = false;

// Utility Functions
function hideAlert() {
  if (divAlert) divAlert.style.display = 'none';
}

function showAlert(msg: string, type: string = 'success') {
  let cls = ALERT_SUCCESS_CLASS;

  if (divAlert) {
    divAlert.innerHTML = msg;
    divAlert.style.display = 'block';

    if (type === 'error') {
      divAlert.classList.remove(ALERT_SUCCESS_CLASS);
      cls = ALERT_ERROR_CLASS;
    } else {
      divAlert.classList.remove(ALERT_ERROR_CLASS);
    }

    divAlert.classList.add(cls);
  }
}

function updateBtnConvert() {
  if (isLoading) {
    btnConvert.textContent = 'Converting...';
  } else {
    btnConvert.textContent = 'CONVERT';
  }
}

async function processFiles(files: FileList) {
  try {
    const emBuffer = await convert(files);

    showAlert('Conversion successful. Download will start automatically.');
    isLoading = false;
    updateBtnConvert();

    const blob = new Blob([emBuffer]);

    const filename = `em-${createTimestamp()}.xlsx`;
    const blobUrl = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = blobUrl;
    a.download = filename;

    document.body.appendChild(a);
    a.click();

    document.body.removeChild(a);
    URL.revokeObjectURL(blobUrl);
  } catch (error) {
    const msg =
      'ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and that the layout was not altered.';
    showAlert(msg, 'error');
    isLoading = false;
    console.error(error);
    updateBtnConvert();
  }
}

function handleSubmit(event: SubmitEvent) {
  event.preventDefault();

  hideAlert();
  isLoading = true;
  updateBtnConvert();
  const { files } = fileInput;

  if (!files) throw new Error('Missing file(s)!');

  processFiles(files).catch(e => console.log(e));
}

// Initialization
hideAlert();
updateBtnConvert();

// Event Listeners
excelForm?.addEventListener('submit', handleSubmit);
