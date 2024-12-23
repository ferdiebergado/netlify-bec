import './sass/app.scss';
import { config } from './config.js';
import { ExpenditureMatrix } from './expenditureMatrix.js';
import type { ExcelFile, ExpenditureFile } from './types/globals.js';
import { createTimestamp } from './utils.js';
import { BudgetEstimateParseError } from './parseError.js';

// Constants
const ALERT_SUCCESS_CLASS = 'alert-success';
const ALERT_ERROR_CLASS = 'alert-error';

// Elements
const excelForm = document.forms.namedItem('excelForm');
const emFileInput = document.getElementById('emFile') as HTMLInputElement;
const beFilesInput = document.getElementById(
  'budgetEstimates',
) as HTMLInputElement;
const divAlert = document.getElementById('alert') as HTMLDivElement;
const btnConvert = document.getElementById('convert') as HTMLButtonElement;
const dropContainer = document.getElementById(
  'dropcontainer',
) as HTMLLabelElement;
const overlay = document.getElementById('overlay') as HTMLDivElement;

let isLoading = false;

function hideAlert() {
  divAlert.style.display = 'none';
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

function updateConvertBtn() {
  if (isLoading) {
    btnConvert.textContent = 'Converting...';
    overlay.style.display = 'block';
  } else {
    btnConvert.textContent = 'CONVERT';
    overlay.style.display = 'none';
  }
}

function initiateDownload(em: ExpenditureFile) {
  const { programTitle, buffer } = em;
  const blob = new Blob([buffer]);
  const blobUrl = URL.createObjectURL(blob);
  const filename = `Expenditure - ${programTitle || createTimestamp()}.xlsx`;

  const a = document.createElement('a');
  a.href = blobUrl;
  a.download = filename;

  document.body.appendChild(a);
  a.click();

  document.body.removeChild(a);
  URL.revokeObjectURL(blobUrl);
}

async function processFiles(filelist: FileList): Promise<ExpenditureFile> {
  const emError = 'em is required';
  const { files } = emFileInput;

  if (!files) throw new Error(emError);

  const emFile = files[0];

  if (!emFile) throw new Error(emError);

  const emArrayBuffer = await emFile.arrayBuffer();
  const em: ExcelFile = {
    filename: emFile.name,
    buffer: emArrayBuffer,
  };
  const expenditureMatrix =
    await ExpenditureMatrix.createAsync<ExpenditureMatrix>(em);

  const budgetEstimates: ExcelFile[] = await Promise.all(
    [...filelist].map(async file => ({
      filename: file.name,
      buffer: await file.arrayBuffer(),
    })),
  );

  const output = await expenditureMatrix.fromBudgetEstimates(budgetEstimates);
  return output;
}

function handleSubmit(event: SubmitEvent) {
  event.preventDefault();

  hideAlert();
  isLoading = true;
  updateConvertBtn();

  const { files } = beFilesInput;

  if (!files || files.length === 0) throw new Error('No files provided!');

  processFiles(files)
    .then(converted => {
      if (converted) {
        initiateDownload(converted);
        showAlert('Conversion successful. Download will start automatically.');
      }
    })
    .catch((error: unknown) => {
      if (error instanceof Error) {
        handleError(error);
      } else {
        console.error('Unknown error occurred:', error);
        showAlert('An unknown error occurred.', 'error');
      }
    })
    .finally(() => {
      isLoading = false;
      updateConvertBtn();
    });
}

function handleError(error: Error) {
  console.error(error);

  const header = `<p><b>ERROR:</b></p>`;

  let msg: string;

  if (error instanceof BudgetEstimateParseError) {
    msg = `<p>${error.message}</p>`;

    const { activity, file, sheet } = error.details;

    if (activity) msg += `<p><b>Activity Title:</b> ${activity}</p>`;

    msg += `<p><b>File:</b> ${file}</p>`;

    if (sheet) msg += `<p><b>Sheet:</b> ${sheet}</p>`;
  } else {
    msg = `
    <p>An error occurred during conversion. Please check the following and try again:</p>
  <ul>
  <li>You are using the official Budget Estimate <a href="${config.paths.beTemplate}">template</a>.</li>
  <li>The activity details are filled up.</li>
  <li>The layout of the template was not altered.</li>
  </ul>
  `;
  }

  showAlert(header + msg, 'error');
}

// Initialization
hideAlert();
updateConvertBtn();

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
  beFilesInput.files = e.dataTransfer?.files || null;
});
