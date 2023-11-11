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

let isLoading = false;

const hideAlert = () => {
  if (divAlert) {
    divAlert.style.display = 'none';
  }
};

const showAlert = (msg: string, type: string = 'success') => {
  if (divAlert) {
    divAlert.innerHTML = msg;
    divAlert.style.display = 'block';
    const clsSuccess = 'alert-success';
    const clsError = 'alert-error';
    let cls = clsSuccess;

    if (type === 'error') {
      divAlert.classList.remove(clsSuccess);
      cls = 'alert-error';
    } else {
      divAlert.classList.remove(clsError);
    }

    divAlert.classList.add(cls);
  }
};

const toggleSpinner = (el: HTMLElement | null, val: string) => {
  if (el) {
    if (isLoading) {
      el.setAttribute('aria-busy', 'true');
    } else {
      el.removeAttribute('aria-busy');
    }
    el.textContent = val;
  }
};

if (!excelForm || !fileInput) {
  throw new Error('Form or file input not found.');
  // Handle case where form or file input is not found
}

hideAlert();

toggleSpinner(btnConvert, 'Convert');

excelForm.addEventListener('submit', async event => {
  event.preventDefault();

  hideAlert();
  isLoading = true;
  toggleSpinner(btnConvert, 'Converting...');

  const formData = new FormData();
  const selectedFile = fileInput.files?.[0];

  if (!selectedFile) {
    throw new Error('No file selected for conversion.');
  }

  formData.append('excelFile', selectedFile);

  try {
    const res = await fetch('/.netlify/functions/convert', {
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
    const msg = `ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and the layout was not changed.`;
    console.error(msg, error);
    // Handle error as needed
    showAlert(msg, 'error');
    isLoading = false;
    toggleSpinner(btnConvert, 'Convert');
  }
});
