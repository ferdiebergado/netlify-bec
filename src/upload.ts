// <label for="images" class="drop-container" id="dropcontainer">
//   <span class="drop-title">Drop files here</span>
//   or
//   <input type="file" id="images" accept="image/*" multiple required>
// </label>

const dropContainer = document.getElementById('dropcontainer');
const fileInput = document.getElementById('images');

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
