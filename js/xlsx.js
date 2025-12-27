// XLSX Converter JavaScript
const state = {
    selectedType: null,
    selectedFile: null,
    selectedFiles: [],
    downloadBlob: null,
}
const API_BASE = 'http://localhost:3000'


// Mobile menu toggle
const mobileMenuBtn = document.getElementById('mobileMenuBtn')
const navLinks = document.getElementById('navLinks')

if (mobileMenuBtn) {
    mobileMenuBtn.addEventListener('click', () => {
        navLinks.classList.toggle('show')
    })
}

// Hero button
const startBtn = document.getElementById('startBtn')
const heroSection = document.getElementById('heroSection')
const mainContainer = document.getElementById('mainContainer')

if (startBtn) {
    startBtn.addEventListener('click', () => {
        heroSection.style.display = 'none'
        mainContainer.style.display = 'block'
    })
}

// DOM elements
const typeCards = document.querySelectorAll('.type-card')
const uploadSection = document.getElementById('uploadSection')
const selectedTypeTitle = document.getElementById('selectedTypeTitle')
const fileInput = document.getElementById('fileInput')
const uploadBox = document.getElementById('uploadBox')
const fileInfo = document.getElementById('fileInfo')
const fileName = document.getElementById('fileName')
const fileSize = document.getElementById('fileSize')
const removeFile = document.getElementById('removeFile')
const convertBtn = document.getElementById('convertBtn')
const backBtn = document.getElementById('backBtn')
const progressSection = document.getElementById('progressSection')
const progressFill = document.getElementById('progressFill')
const progressText = document.getElementById('progressText')
const resultSection = document.getElementById('resultSection')
const resultSuccess = document.getElementById('resultSuccess')
const resultError = document.getElementById('resultError')
const resultMessage = document.getElementById('resultMessage')
const errorMessage = document.getElementById('errorMessage')
const downloadBtn = document.getElementById('downloadBtn')
const convertAnotherBtn = document.getElementById('convertAnotherBtn')
const singleUpload = document.getElementById('singleUpload')
const multipleUpload = document.getElementById('multipleUpload')
const multipleFileInput = document.getElementById('multipleFileInput')
const multipleUploadBox = document.getElementById('multipleUploadBox')
const filesList = document.getElementById('filesList')
const conversionOptions = document.getElementById('conversionOptions')
const sheetsOption = document.getElementById('sheetsOption')
const sheetsInput = document.getElementById('sheetsInput')
const metadataResult = document.getElementById('metadataResult')
const metadataGrid = document.getElementById('metadataGrid')
const sheetsInfo = document.getElementById('sheetsInfo')
const sheetsList = document.getElementById('sheetsList')

// Event listeners
typeCards.forEach(card => {
    card.addEventListener('click', () => selectType(card.dataset.type))
})

uploadBox.addEventListener('click', () => fileInput.click())
fileInput.addEventListener('change', handleFileSelect)
removeFile.addEventListener('click', clearFile)
backBtn.addEventListener('click', goBack)
convertBtn.addEventListener('click', startConversion)
convertAnotherBtn.addEventListener('click', resetAll)

multipleUploadBox.addEventListener('click', () => multipleFileInput.click())
multipleFileInput.addEventListener('change', handleMultipleFiles)

// Drag & Drop
uploadBox.addEventListener('dragover', e => {
    e.preventDefault()
    uploadBox.classList.add('drag-over')
})

uploadBox.addEventListener('dragleave', () => {
    uploadBox.classList.remove('drag-over')
})

uploadBox.addEventListener('drop', e => {
    e.preventDefault()
    uploadBox.classList.remove('drag-over')
    const files = e.dataTransfer.files
    if (files.length > 0) {
        fileInput.files = files
        handleFileSelect()
    }
})

// Select conversion type
function selectType(type) {
    state.selectedType = type
    
    const titles = {
        'xlsx-to-pdf': '📄 XLSX → PDF Konvertatsiya',
        'xlsx-to-images': '🖼️ XLSX → Images Konvertatsiya',
        'xlsx-sheets-to-pdf': '📑 XLSX Sheetlar → PDF',
        'xlsx-merge-pdf': '🔗 Ko\'p XLSX → PDF Birlashtirish',
        'xlsx-metadata': '🔍 XLSX Metadata Tahlili',
        'xlsx-to-csv': '📋 XLSX → CSV Konvertatsiya',
    }
    
    selectedTypeTitle.textContent = titles[type]
    
    // Show appropriate upload section
    if (type === 'xlsx-merge-pdf') {
        singleUpload.style.display = 'none'
        multipleUpload.style.display = 'block'
    } else {
        singleUpload.style.display = 'block'
        multipleUpload.style.display = 'none'
    }
    
    // Show options
    conversionOptions.style.display = 'none'
    sheetsOption.style.display = 'none'
    
    if (type === 'xlsx-sheets-to-pdf') {
        conversionOptions.style.display = 'block'
        sheetsOption.style.display = 'block'
    }
    
    document.querySelector('.conversion-types').style.display = 'none'
    uploadSection.style.display = 'block'
}

// Handle single file
function handleFileSelect() {
    const file = fileInput.files[0]
    if (!file) return
    
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        alert('Faqat XLSX/XLS fayllar qabul qilinadi!')
        return
    }
    
    if (file.size > 20 * 1024 * 1024) {
        alert('Fayl hajmi 20MB dan oshmasligi kerak!')
        return
    }
    
    state.selectedFile = file
    
    fileName.textContent = file.name
    fileSize.textContent = formatFileSize(file.size)
    
    uploadBox.style.display = 'none'
    fileInfo.style.display = 'flex'
    convertBtn.disabled = false
}

// Handle multiple files
function handleMultipleFiles() {
    const files = Array.from(multipleFileInput.files)
    
    if (files.length === 0) return
    
    if (files.length > 10) {
        alert('Maksimal 10 ta fayl!')
        return
    }
    
    const validFiles = files.filter(file => {
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            alert(`${file.name} - Faqat XLSX/XLS!`)
            return false
        }
        if (file.size > 20 * 1024 * 1024) {
            alert(`${file.name} - 20MB dan katta!`)
            return false
        }
        return true
    })
    
    state.selectedFiles = validFiles
    displayFilesList()
    convertBtn.disabled = validFiles.length === 0
}

// Display files list
function displayFilesList() {
    filesList.innerHTML = ''
    
    state.selectedFiles.forEach((file, index) => {
        const item = document.createElement('div')
        item.className = 'file-item'
        item.innerHTML = `
            <div class="file-details">
                <span class="file-name">${file.name}</span>
                <span class="file-size">${formatFileSize(file.size)}</span>
            </div>
            <button class="btn-remove" onclick="removeFileFromList(${index})">✕</button>
        `
        filesList.appendChild(item)
    })
}

// Remove file from list
window.removeFileFromList = function(index) {
    state.selectedFiles.splice(index, 1)
    displayFilesList()
    convertBtn.disabled = state.selectedFiles.length === 0
}

// Clear single file
function clearFile() {
    state.selectedFile = null
    fileInput.value = ''
    fileInfo.style.display = 'none'
    uploadBox.style.display = 'block'
    convertBtn.disabled = true
}

// Start conversion
async function startConversion() {
    const endpoints = {
        'xlsx-to-pdf': 'http://localhost:3000/api/convert/xlsx-to-pdf',
        'xlsx-to-images': 'http://localhost:3000/api/convert/xlsx-to-images',
        'xlsx-sheets-to-pdf': 'http://localhost:3000/api/convert/xlsx-sheets-to-pdf',
        'xlsx-merge-pdf': 'http://localhost:3000/api/convert/xlsx-merge-to-pdf',
        'xlsx-metadata': 'http://localhost:3000/api/analyze/xlsx-metadata',
        'xlsx-to-csv': 'http://localhost:3000/api/convert/xlsx-to-csv',
    }
    

    const endpoint = endpoints[state.selectedType]

    if (!endpoint) {
        alert('Noto‘g‘ri konvertatsiya turi!')
        return
    }

    uploadSection.style.display = 'none'
    progressSection.style.display = 'block'
    updateProgress(10, 'Fayl yuklanmoqda...')

    try {
        const formData = new FormData()

        // Fayllarni qo‘shish
        if (state.selectedType === 'xlsx-merge-pdf') {
            if (!state.selectedFiles.length) throw new Error('Hech qanday fayl tanlanmadi!')
            state.selectedFiles.forEach(file => formData.append('files[]', file))
        } else {
            if (!state.selectedFile) throw new Error('Fayl tanlanmadi!')
            formData.append('file', state.selectedFile)
        }

        // Agar sheets tanlanadigan konvertatsiya bo‘lsa
        if (state.selectedType === 'xlsx-sheets-to-pdf') {
            formData.append('sheets', sheetsInput.value || '')
        }

        updateProgress(40, 'Serverga yuborilmoqda...')

        const response = await fetch(endpoint, {
            method: 'POST',
            body: formData,
        })

        const contentType = response.headers.get('content-type') || ''

        // ❌ HTTP xato
        if (!response.ok) {
            let errorMessage = 'Server xatosi'
            if (contentType.includes('application/json')) {
                const err = await response.json()
                errorMessage = err.error || errorMessage
            } else {
                const text = await response.text()
                errorMessage = text || errorMessage
            }
            throw new Error(errorMessage)
        }

        updateProgress(80, 'Natija tayyorlanmoqda...')

        // ✅ METADATA
        if (state.selectedType === 'xlsx-metadata') {
            const data = await response.json()
            showMetadataResult(data)
            return
        }

        // ✅ FILE / ZIP / PDF
        const blob = await response.blob()
        state.downloadBlob = blob

        updateProgress(100, 'Tayyor!')

        setTimeout(() => {
            progressSection.style.display = 'none'
            showSuccess('Fayl muvaffaqiyatli konvertatsiya qilindi!')
        }, 400)

    } catch (error) {
        console.error('Konvertatsiya xatosi:', error)
        progressSection.style.display = 'none'
        showError(error.message || 'Konvertatsiya xatosi')
    }
}



// Update progress
function updateProgress(percent, text) {
    progressFill.style.width = `${percent}%`
    progressText.textContent = text
}

// Show success
function showSuccess(message) {
    resultSection.style.display = 'block'
    resultSuccess.style.display = 'block'
    resultError.style.display = 'none'
    metadataResult.style.display = 'none'
    resultMessage.textContent = message
}

// Show error
function showError(message) {
    resultSection.style.display = 'block'
    resultSuccess.style.display = 'none'
    resultError.style.display = 'block'
    metadataResult.style.display = 'none'
    errorMessage.textContent = message
}

// Show metadata
function showMetadataResult(data) {
    progressSection.style.display = 'none'
    resultSection.style.display = 'block'
    resultSuccess.style.display = 'none'
    resultError.style.display = 'none'
    metadataResult.style.display = 'block'
    
    metadataGrid.innerHTML = ''
    
    const fields = {
        'fileName': 'Fayl nomi',
        'fileSizeFormatted': 'Fayl hajmi',
        'totalSheets': 'Jami sheetlar',
    }
    
    Object.entries(fields).forEach(([key, label]) => {
        if (data[key]) {
            const item = document.createElement('div')
            item.className = 'metadata-item'
            item.innerHTML = `
                <div class="metadata-label">${label}</div>
                <div class="metadata-value">${data[key]}</div>
            `
            metadataGrid.appendChild(item)
        }
    })
    
    // Sheets info
    if (data.sheets && data.sheets.length > 0) {
        sheetsInfo.style.display = 'block'
        sheetsList.innerHTML = ''
        
        data.sheets.forEach((sheet, index) => {
            const item = document.createElement('div')
            item.className = 'sheet-item'
            item.innerHTML = `
                <div class="sheet-name">📄 ${index}: ${sheet.name}</div>
                <div class="sheet-stats">
                    ${sheet.rows} qator × ${sheet.columns} ustun = ${sheet.cells} katak
                </div>
            `
            sheetsList.appendChild(item)
        })
    }
}

// Download file
downloadBtn.addEventListener('click', () => {
    if (!state.downloadBlob) return
    
    const extensions = {
        'xlsx-to-pdf': 'pdf',
        'xlsx-to-images': 'zip',
        'xlsx-sheets-to-pdf': 'pdf',
        'xlsx-merge-pdf': 'pdf',
        'xlsx-to-csv': 'zip',
    }
    
    const ext = extensions[state.selectedType] || 'pdf'
    const filename = `excel-converted-${Date.now()}.${ext}`
    
    const url = URL.createObjectURL(state.downloadBlob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
})

// Go back
function goBack() {
    uploadSection.style.display = 'none'
    document.querySelector('.conversion-types').style.display = 'block'
    clearFile()
    state.selectedFiles = []
    filesList.innerHTML = ''
}

// Reset all
function resetAll() {
    resultSection.style.display = 'none'
    uploadSection.style.display = 'none'
    document.querySelector('.conversion-types').style.display = 'block'
    clearFile()
    state.selectedFiles = []
    state.downloadBlob = null
    filesList.innerHTML = ''
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
}