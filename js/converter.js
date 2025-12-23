
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

