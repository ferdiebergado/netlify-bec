import { API_ENDPOINT } from '../server/constants'
import { createTimestamp } from '../server/utils'

// Constants
const ALERT_SUCCESS_CLASS = 'alert-success'
const ALERT_ERROR_CLASS = 'alert-error'
const LOADING_CLASS = 'aria-busy'

// Elements
const excelForm = document.getElementById('excelForm') as HTMLFormElement | null
const fileInput = document.getElementById(
  'excelFile',
) as HTMLInputElement | null
const divAlert = document.getElementById('alert') as HTMLDivElement | null
const btnConvert = document.getElementById(
  'convert',
) as HTMLButtonElement | null

// Check if required elements are present
if (!excelForm || !fileInput || !divAlert || !btnConvert) {
  throw new Error('One or more required elements not found.')
}

let isLoading = false

// Utility Functions
function hideAlert() {
  divAlert!.style.display = 'none'
}

function showAlert(msg: string, type: string = 'success') {
  let cls = ALERT_SUCCESS_CLASS

  divAlert!.innerHTML = msg
  divAlert!.style.display = 'block'

  if (type === 'error') {
    divAlert!.classList.remove(ALERT_SUCCESS_CLASS)
    cls = ALERT_ERROR_CLASS
  } else {
    divAlert!.classList.remove(ALERT_ERROR_CLASS)
  }

  divAlert!.classList.add(cls)
}

function toggleSpinner(el: HTMLElement, val: string) {
  if (isLoading) {
    el.setAttribute(LOADING_CLASS, 'true')
  } else {
    el.removeAttribute(LOADING_CLASS)
  }

  // eslint-disable-next-line no-param-reassign
  el.textContent = val
}

async function handleConversion(selectedFiles: FileList) {
  const formData = new FormData()

  ;[...selectedFiles].forEach(file => {
    formData.append('excelFile', file)
  })

  try {
    const res = await fetch(API_ENDPOINT, {
      method: 'POST',
      body: formData,
    })

    if (!res.ok) {
      throw new Error('Conversion failed!')
    }

    showAlert('Conversion successful. Download will start automatically.')
    isLoading = false
    toggleSpinner(btnConvert!, 'Convert')

    const blob = await res.blob()
    const contentDisposition = res.headers.get('Content-Disposition')
    const filenameMatch =
      contentDisposition && contentDisposition.match(/filename="(.+?)"/)
    const filename = filenameMatch
      ? filenameMatch[1]
      : `em-${createTimestamp()}.xlsx`
    const blobUrl = URL.createObjectURL(blob)

    const a = document.createElement('a')
    a.href = blobUrl
    a.download = filename

    document.body.appendChild(a)
    a.click()

    document.body.removeChild(a)
    URL.revokeObjectURL(blobUrl)
  } catch (error) {
    const msg =
      'ERROR:<br>An error occurred during conversion.<br>Please make sure that you are using the official Budget Estimate template and that the layout was not altered.'
    showAlert(msg, 'error')
    isLoading = false
    toggleSpinner(btnConvert!, 'Convert')
  }
}

function handleSubmit(event: SubmitEvent) {
  event.preventDefault()

  hideAlert()
  isLoading = true
  toggleSpinner(btnConvert!, 'Converting...')

  const selectedFiles = fileInput!.files

  if (!selectedFiles) {
    throw new Error('No file selected for conversion.')
  }

  handleConversion(selectedFiles).catch(e => console.log(e))
}

// Initialization
hideAlert()
toggleSpinner(btnConvert, 'Convert')

// Event Listeners
excelForm.addEventListener('submit', handleSubmit)
