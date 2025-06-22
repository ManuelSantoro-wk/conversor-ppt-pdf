let dropArea = document.getElementById('drop-area');
let fileInput = document.getElementById('fileElem');

// Prevenir comportamentos padrão
;['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => e.preventDefault(), false);
    document.body.addEventListener(eventName, e => e.preventDefault(), false);
});

// Highlight visual
;['dragenter', 'dragover'].forEach(eventName => {
    dropArea.addEventListener(eventName, () => dropArea.classList.add('highlight'), false);
});

;['dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, () => dropArea.classList.remove('highlight'), false);
});

// Handle drop
dropArea.addEventListener('drop', e => {
    if (e.dataTransfer.files.length > 0) {
        fileInput.files = e.dataTransfer.files;
    }
});

