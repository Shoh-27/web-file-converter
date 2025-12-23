
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

let selectedFiles = [];
let currentConverter = '';
let pdfPages = [];

const converters = {
    'txt-pdf': { title: 'TXT → PDF', accept: '.txt', multiple: false },
    'img-pdf': { title: 'Image → PDF', accept: 'image/*', multiple: false },
    'pdf-txt': { title: 'PDF → TXT', accept: '.pdf', multiple: false },
    'docx-pdf': { title: 'DOCX → PDF', accept: '.doc,.docx', multiple: false },
    'multi-img-pdf': { title: 'Multiple Images → PDF', accept: 'image/*', multiple: true },
    'pdf-split': { title: 'PDF Split', accept: '.pdf', multiple: false },
    'pdf-merge': { title: 'PDF Merge', accept: '.pdf', multiple: true },
    'file-analyzer': { title: 'File Size Analyzer', accept: '*', multiple: false }
};

function openConverter(type) {
    currentConverter = type;
    const converter = converters[type];
    document.getElementById('modalTitle').textContent = converter.title;
    document.getElementById('fileInput').accept = converter.accept;
    document.getElementById('fileInput').multiple = converter.multiple;

    selectedFiles = [];
    pdfPages = [];
    updateFileList();
    document.getElementById('converterModal').classList.add('active');
    document.getElementById('resultArea').style.display = 'none';
    document.getElementById('progressBar').style.display = 'none';
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('optionsArea').style.display = 'none';
    document.getElementById('optionsArea').innerHTML = '';
}

function closeModal() {
    document.getElementById('converterModal').classList.remove('active');
}

const uploadArea = document.getElementById('uploadArea');

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = Array.from(e.dataTransfer.files);
    addFiles(files);
});

document.getElementById('fileInput').addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    addFiles(files);
});
