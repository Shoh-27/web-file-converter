
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

function addFiles(files) {
    if (converters[currentConverter].multiple) {
        selectedFiles = selectedFiles.concat(files);
    } else {
        selectedFiles = files.slice(0, 1);
    }
    updateFileList();

    if (currentConverter === 'pdf-split') {
        loadPdfPages(files[0]);
    }
}

function updateFileList() {
    const fileList = document.getElementById('fileList');
    fileList.innerHTML = '';

    selectedFiles.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
                    <span>${file.name} (${formatFileSize(file.size)})</span>
                    <button onclick="removeFile(${index})">O'chirish</button>
                `;
        fileList.appendChild(item);
    });

    document.getElementById('convertBtn').disabled = selectedFiles.length === 0;
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    updateFileList();
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

async function loadPdfPages(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

    pdfPages = [];
    const optionsArea = document.getElementById('optionsArea');
    optionsArea.style.display = 'block';
    optionsArea.innerHTML = '<h3>Sahifalarni tanlang:</h3><div class="page-preview" id="pagePreview"></div>';

    const pagePreview = document.getElementById('pagePreview');

    for (let i = 1; i <= pdf.numPages; i++) {
        pdfPages.push(i);
        const pageDiv = document.createElement('div');
        pageDiv.className = 'page-item selected';
        pageDiv.innerHTML = `<p>Sahifa ${i}</p>`;
        pageDiv.onclick = () => togglePage(i, pageDiv);
        pagePreview.appendChild(pageDiv);
    }
}

function togglePage(pageNum, element) {
    const index = pdfPages.indexOf(pageNum);
    if (index > -1) {
        pdfPages.splice(index, 1);
        element.classList.remove('selected');
    } else {
        pdfPages.push(pageNum);
        element.classList.add('selected');
    }
    pdfPages.sort((a, b) => a - b);
}

async function convertFiles() {
    const progressBar = document.getElementById('progressBar');
    const progressFill = document.getElementById('progressFill');
    const resultArea = document.getElementById('resultArea');
    const fileInfo = document.getElementById('fileInfo');

    progressBar.style.display = 'block';
    resultArea.style.display = 'none';
    fileInfo.style.display = 'none';

    for (let i = 0; i <= 100; i += 10) {
        progressFill.style.width = i + '%';
        progressFill.textContent = i + '%';
        await new Promise(resolve => setTimeout(resolve, 100));
    }

    try {
        switch (currentConverter) {
            case 'txt-pdf':
                await convertTxtToPdf();
                break;
            case 'img-pdf':
            case 'multi-img-pdf':
                await convertImageToPdf();
                break;
            case 'pdf-txt':
                await convertPdfToTxt();
                break;
            case 'docx-pdf':
                await convertDocxToPdf();
                break;
            case 'pdf-split':
                await splitPdf();
                break;
            case 'pdf-merge':
                await mergePdf();
                break;
            case 'file-analyzer':
                analyzeFile();
                break;
        }
    } catch (error) {
        alert('Xatolik yuz berdi: ' + error.message);
        progressBar.style.display = 'none';
    }
}

async function convertTxtToPdf() {
    const file = selectedFiles[0];
    const text = await file.text();

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    const lines = pdf.splitTextToSize(text, 180);
    let y = 20;

    lines.forEach(line => {
        if (y > 280) {
            pdf.addPage();
            y = 20;
        }
        pdf.text(line, 15, y);
        y += 7;
    });

    const pdfBlob = pdf.output('blob');
    showResult([{ blob: pdfBlob, name: file.name.replace('.txt', '.pdf') }]);
}

async function convertImageToPdf() {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        const img = await readFileAsDataURL(file);

        if (i > 0) pdf.addPage();

        const imgProps = pdf.getImageProperties(img);
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;

        pdf.addImage(img, 'JPEG', 0, 0, pdfWidth, pdfHeight);
    }

    const pdfBlob = pdf.output('blob');
    showResult([{ blob: pdfBlob, name: 'images.pdf' }]);
}

async function convertPdfToTxt() {
    const file = selectedFiles[0];
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

    let fullText = '';

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str).join(' ');
        fullText += `--- Sahifa ${i} ---\n${pageText}\n\n`;
    }

    const txtBlob = new Blob([fullText], { type: 'text/plain' });
    showResult([{ blob: txtBlob, name: file.name.replace('.pdf', '.txt') }]);
}

async function convertDocxToPdf() {
    const file = selectedFiles[0];
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    const lines = pdf.splitTextToSize(result.value, 180);
    let y = 20;

    lines.forEach(line => {
        if (y > 280) {
            pdf.addPage();
            y = 20;
        }
        pdf.text(line, 15, y);
        y += 7;
    });

    const pdfBlob = pdf.output('blob');
    showResult([{ blob: pdfBlob, name: file.name.replace(/\.(doc|docx)$/, '.pdf') }]);
}

async function splitPdf() {
    const file = selectedFiles[0];
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

    const { jsPDF } = window.jspdf;
    const results = [];

    for (const pageNum of pdfPages) {
        const page = await pdf.getPage(pageNum);
        const viewport = page.getViewport({ scale: 2.0 });

        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.width = viewport.width;
        canvas.height = viewport.height;

        await page.render({ canvasContext: context, viewport: viewport }).promise;

        const imgData = canvas.toDataURL('image/jpeg', 1.0);
        const newPdf = new jsPDF({
            orientation: viewport.width > viewport.height ? 'landscape' : 'portrait',
            unit: 'px',
            format: [viewport.width, viewport.height]
        });

        newPdf.addImage(imgData, 'JPEG', 0, 0, viewport.width, viewport.height);

        results.push({
            blob: newPdf.output('blob'),
            name: `${file.name.replace('.pdf', '')}_page_${pageNum}.pdf`
        });
    }

    showResult(results);
}

async function mergePdf() {
    const { jsPDF } = window.jspdf;
    const mergedPdf = new jsPDF();
    let firstPage = true;

    for (const file of selectedFiles) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const viewport = page.getViewport({ scale: 2.0 });

            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');
            canvas.width = viewport.width;
            canvas.height = viewport.height;

            await page.render({ canvasContext: context, viewport: viewport }).promise;

            const imgData = canvas.toDataURL('image/jpeg', 1.0);

            if (!firstPage) {
                mergedPdf.addPage();
            }
            firstPage = false;

            const pdfWidth = mergedPdf.internal.pageSize.getWidth();
            const pdfHeight = (viewport.height * pdfWidth) / viewport.width;

            mergedPdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
        }
    }

    const pdfBlob = mergedPdf.output('blob');
    showResult([{ blob: pdfBlob, name: 'merged.pdf' }]);
}

function analyzeFile() {
    const file = selectedFiles[0];
    const fileInfo = document.getElementById('fileInfo');

    fileInfo.innerHTML = `
                <h3>📊 Fayl ma'lumotlari:</h3>
                <p><strong>Nomi:</strong> ${file.name}</p>
                <p><strong>Turi:</strong> ${file.type || 'Noma\'lum'}</p>
                <p><strong>Hajmi:</strong> ${formatFileSize(file.size)}</p>
                <p><strong>O'zgartirilgan:</strong> ${new Date(file.lastModified).toLocaleString('uz-UZ')}</p>
            `;

    fileInfo.style.display = 'block';
    document.getElementById('resultArea').style.display = 'block';
    document.getElementById('resultMessage').textContent = 'Fayl tahlili yakunlandi!';
}

function readFileAsDataURL(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.readAsDataURL(file);
    });
}

function showResult(results) {
    const resultArea = document.getElementById('resultArea');
    const downloadArea = document.getElementById('downloadArea');
    downloadArea.innerHTML = '';

    results.forEach((result, index) => {
        const url = URL.createObjectURL(result.blob);
        const btn = document.createElement('a');
        btn.className = 'download-btn';
        btn.href = url;
        btn.download = result.name;
        btn.textContent = `📥 ${result.name}`;
        downloadArea.appendChild(btn);
    });

    document.getElementById('resultMessage').textContent =
        results.length > 1
            ? `${results.length} ta fayl tayyor!`
            : 'Faylingiz tayyor!';
    resultArea.style.display = 'block';
}


<!DOCTYPE html>
<html lang="uz">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>File Converter</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 Universal File Converter</h1>
            <p>Fayl konvertatsiyasi oson va tez!</p>
        </div>

        <div class="converter-grid">
            <div class="converter-card" onclick="openConverter('txt-pdf')">
                <h3>1️⃣ TXT → PDF</h3>
                <p>Text fayllarni PDF formatga o'tkazing</p>
            </div>

            <div class="converter-card" onclick="openConverter('img-pdf')">
                <h3>2️⃣ Image → PDF</h3>
                <p>Rasmlarni PDF formatga o'tkazing</p>
            </div>

            <div class="converter-card" onclick="openConverter('pdf-txt')">
                <h3>3️⃣ PDF → TXT</h3>
                <p>PDF dan matn ajratib oling</p>
            </div>

            <div class="converter-card" onclick="openConverter('docx-pdf')">
                <h3>4️⃣ DOCX → PDF</h3>
                <p>Word hujjatlarni PDF ga aylantiring</p>
            </div>

            <div class="converter-card" onclick="openConverter('multi-img-pdf')">
                <h3>5️⃣ Multiple Images → PDF</h3>
                <p>Ko'p rasmlarni bitta PDF ga birlashtiring</p>
            </div>

            <div class="converter-card" onclick="openConverter('pdf-split')">
                <h3>6️⃣ PDF Split</h3>
                <p>PDF ni alohida sahifalarga bo'ling</p>
            </div>

            <div class="converter-card" onclick="openConverter('pdf-merge')">
                <h3>7️⃣ PDF Merge</h3>
                <p>Bir nechta PDF ni birlashtiring</p>
            </div>

            <div class="converter-card" onclick="openConverter('file-analyzer')">
                <h3>🛠️ File Size Analyzer</h3>
                <p>Fayl hajmini tahlil qiling</p>
            </div>
        </div>
    </div>

    <div class="modal" id="converterModal">
        <div class="modal-content">
            <span class="modal-close" onclick="closeModal()">&times;</span>
            <h2 id="modalTitle">Converter</h2>
            
            <div class="upload-area" id="uploadArea">
                <p>📁 Faylni shu yerga tashlang yoki tanlang</p>
                <input type="file" id="fileInput" multiple>
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()">
                    Fayl tanlash
                </button>
            </div>

            <div class="file-list" id="fileList"></div>

            <div class="options-area" id="optionsArea" style="display:none;"></div>

            <button class="convert-btn" id="convertBtn" onclick="convertFiles()" disabled>
                Konvertatsiya qilish
            </button>

            <div class="progress-bar" id="progressBar">
                <div class="progress-fill" id="progressFill">0%</div>
            </div>

            <div class="result-area" id="resultArea">
                <h3>✅ Tayyor!</h3>
                <p id="resultMessage"></p>
                <div id="downloadArea"></div>
            </div>

            <div class="file-info" id="fileInfo" style="display:none;"></div>
        </div>
    </div>

    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
    <script src="converter.js"></script>
</body>
</html>
