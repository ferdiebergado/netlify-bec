import config from './config';
import convert from './converter';
import { ExpenditureMatrix } from './expenditureMatrix';
import { ExcelFile } from './types/globals';
import { createTimestamp } from './utils';

// Constants
const ALERT_SUCCESS_CLASS = 'alert-success';
const ALERT_ERROR_CLASS = 'alert-error';

// Elements
const excelForm = document.forms.namedItem('excelForm');
const fileInput = document.getElementById('excelFile') as HTMLInputElement;
const divAlert = document.getElementById('alert') as HTMLDivElement;
const btnConvert = document.getElementById('convert') as HTMLButtonElement;
const dropContainer = document.getElementById(
  'dropcontainer',
) as HTMLLabelElement;

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

function initiateDownload(buffer: ArrayBuffer) {
  const blob = new Blob([buffer]);

  const filename = `em-${createTimestamp()}.xlsx`;
  const blobUrl = URL.createObjectURL(blob);

  const a = document.createElement('a');
  a.href = blobUrl;
  a.download = filename;

  document.body.appendChild(a);
  a.click();

  document.body.removeChild(a);
  URL.revokeObjectURL(blobUrl);
}

async function processFiles(files: FileList) {
  try {
    const emTemplate = config.paths.emTemplate;
    const res = await fetch(emTemplate);
    const arrayBuffer = await res.arrayBuffer();
    const expenditureMatrix =
      await ExpenditureMatrix.createAsync<ExpenditureMatrix>(arrayBuffer);

    const excelFiles: ExcelFile[] = [];

    await Promise.allSettled(
      [...files].map(async file => {
        excelFiles.push({
          filename: file.name,
          buffer: await file.arrayBuffer(),
        });
      }),
    );

    const buffer = await expenditureMatrix.fromBudgetEstimates(excelFiles);

    showAlert('Conversion successful. Download will start automatically.');
    isLoading = false;
    updateBtnConvert();

    initiateDownload(buffer);
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

  processFiles(files).catch(e => console.error(e));
}

// Initialization
hideAlert();
updateBtnConvert();

// Event Listeners
excelForm?.addEventListener('submit', handleSubmit);

dropContainer.addEventListener(
  'dragover',
  e => {
    // prevent default to allow drop
    e.preventDefault();
  },
  false,
);

dropContainer.addEventListener('dragenter', () => {
  dropContainer.classList.add('drag-active');
});

dropContainer.addEventListener('dragleave', () => {
  dropContainer.classList.remove('drag-active');
});

dropContainer.addEventListener('drop', e => {
  e.preventDefault();
  dropContainer.classList.remove('drag-active');
  fileInput.files = e.dataTransfer.files;
});
