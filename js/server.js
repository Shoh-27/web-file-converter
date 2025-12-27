const express = require('express')
const multer = require('multer')
const path = require('path')
const fs = require('fs').promises
const { exec } = require('child_process')
const { promisify } = require('util')
const archiver = require('archiver')
const sharp = require('sharp')
const PDFDocument = require('pdfkit')

const execPromise = promisify(exec)
const app = express()
const PORT = process.env.PORT || 3000

// Настройка CORS
app.use((req, res, next) => {
	res.header('Access-Control-Allow-Origin', '*')
	res.header('Access-Control-Allow-Methods', 'GET, POST, DELETE')
	res.header('Access-Control-Allow-Headers', 'Content-Type')
	next()
})

// Статические файлы
app.use(express.static('public'))
app.use('/style', express.static(path.join(__dirname, '../style')))
app.use('/js', express.static(path.join(__dirname)))
app.use('/public', express.static(path.join(__dirname, '../public')))

// Настройка multer для загрузки файлов
const storage = multer.diskStorage({
	destination: async (req, file, cb) => {
		const tmpDir = '/tmp/conversions'
		try {
			await fs.mkdir(tmpDir, { recursive: true })
			cb(null, tmpDir)
		} catch (err) {
			cb(err)
		}
	},
	filename: (req, file, cb) => {
		const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9)
		cb(null, uniqueSuffix + '-' + file.originalname)
	},
})

const upload = multer({
	storage,
	limits: { fileSize: 50 * 1024 * 1024 }, // 50MB лимит
})

// Health check endpoint
app.get('/health', (req, res) => {
	res.status(200).json({ status: 'OK', timestamp: new Date().toISOString() })
})

// API endpoint для конвертации DOCX в PDF
app.post(
	'/api/convert/docx-to-pdf',
	upload.single('file'),
	async (req, res) => {
		let inputPath = null
		let outputPath = null

		try {
			if (!req.file) {
				return res.status(400).json({ error: 'Файл не загружен' })
			}

			inputPath = req.file.path
			const outputDir = path.dirname(inputPath)
			const baseName = path.basename(inputPath, path.extname(inputPath))
			outputPath = path.join(outputDir, `${baseName}.pdf`)

			console.log(`Конвертация: ${inputPath} -> ${outputPath}`)

			// Конвертация через LibreOffice
			const command = `libreoffice --headless --convert-to pdf --outdir "${outputDir}" "${inputPath}"`

			await execPromise(command, {
				timeout: 60000, // 60 секунд таймаут
				maxBuffer: 10 * 1024 * 1024,
			})

			// Проверяем существование выходного файла
			try {
				await fs.access(outputPath)
			} catch {
				throw new Error('PDF файл не был создан')
			}

			// Читаем PDF файл
			const pdfBuffer = await fs.readFile(outputPath)

			// Отправляем файл клиенту
			res.setHeader('Content-Type', 'application/pdf')
			res.setHeader(
				'Content-Disposition',
				`attachment; filename="${baseName}.pdf"`
			)
			res.send(pdfBuffer)
		} catch (error) {
			console.error('Ошибка конвертации:', error)
			res.status(500).json({
				error: 'Ошибка конвертации файла',
				details: error.message,
			})
		} finally {
			// Очистка временных файлов
			try {
				if (inputPath) await fs.unlink(inputPath).catch(() => {})
				if (outputPath) await fs.unlink(outputPath).catch(() => {})
			} catch (err) {
				console.error('Ошибка очистки файлов:', err)
			}
		}
	}
)

app.post('/api/convert/pptx-to-pdf', upload.single('file'), async (req, res) => {
	let inputPath = null
	let outputPath = null

	try {
		if (!req.file) {
			return res.status(400).json({ error: 'Файл не загружен' })
		}

		inputPath = req.file.path
		const outputDir = path.dirname(inputPath)
		const baseName = path.basename(inputPath, path.extname(inputPath))
		outputPath = path.join(outputDir, `${baseName}.pdf`)

		console.log(`PPTX → PDF: ${inputPath} -> ${outputPath}`)

		const command = `libreoffice --headless --convert-to pdf --outdir "${outputDir}" "${inputPath}"`

		await execPromise(command, {
			timeout: 120000,
			maxBuffer: 20 * 1024 * 1024,
		})

		await fs.access(outputPath)
		const pdfBuffer = await fs.readFile(outputPath)

		res.setHeader('Content-Type', 'application/pdf')
		res.setHeader('Content-Disposition', `attachment; filename="${baseName}.pdf"`)
		res.send(pdfBuffer)
	} catch (error) {
		console.error('PPTX → PDF ошибка:', error)
		res.status(500).json({ error: 'Ошибка конвертации', details: error.message })
	} finally {
		if (inputPath) await fs.unlink(inputPath).catch(() => {})
		if (outputPath) await fs.unlink(outputPath).catch(() => {})
	}
})



const XLSX = require('xlsx')
const ExcelJS = require('exceljs')
const PDFDocument = require('pdfkit')
const sharp = require('sharp')
const archiver = require('archiver')

// ==========================================
// 1. XLSX → PDF (Table formatda)
// ==========================================
app.post('/api/convert/xlsx-to-pdf', upload.single('file'), async (req, res) => {
	let inputPath = null
	let outputPath = null

	try {
		if (!req.file) {
			return res.status(400).json({ error: 'Fayl yuklanmadi' })
		}

		inputPath = req.file.path
		outputPath = path.join('/tmp/conversions', `${Date.now()}.pdf`)

		// Excel faylni o'qish
		const workbook = XLSX.readFile(inputPath)
		const sheetName = workbook.SheetNames[0]
		const sheet = workbook.Sheets[sheetName]
		const data = XLSX.utils.sheet_to_json(sheet, { header: 1 })

		// PDF yaratish
		const doc = new PDFDocument({ margin: 30 })
		const stream = require('fs').createWriteStream(outputPath)
		doc.pipe(stream)

		// Sarlavha
		doc.fontSize(16).text(`Sheet: ${sheetName}`, { align: 'center' })
		doc.moveDown()

		// Jadval chizish
		const cellWidth = 80
		const cellHeight = 25
		let yPos = doc.y

		data.forEach((row, rowIndex) => {
			let xPos = 30

			row.forEach((cell, colIndex) => {
				// Border
				doc.rect(xPos, yPos, cellWidth, cellHeight).stroke()

				// Matn
				const text = String(cell || '')
				doc.fontSize(10).text(text, xPos + 5, yPos + 8, {
					width: cellWidth - 10,
					height: cellHeight - 10,
					ellipsis: true,
				})

				xPos += cellWidth
			})

			yPos += cellHeight

			// Yangi sahifa
			if (yPos > 700 && rowIndex < data.length - 1) {
				doc.addPage()
				yPos = 50
			}
		})

		doc.end()

		await new Promise((resolve, reject) => {
			stream.on('finish', resolve)
			stream.on('error', reject)
		})

		const pdfBuffer = await fs.readFile(outputPath)

		res.setHeader('Content-Type', 'application/pdf')
		res.setHeader(
			'Content-Disposition',
			`attachment; filename="excel-${Date.now()}.pdf"`
		)
		res.send(pdfBuffer)
	} catch (error) {
		console.error('XLSX → PDF xato:', error)
		res.status(500).json({ error: 'Konvertatsiya xatosi', details: error.message })
	} finally {
		if (inputPath) await fs.unlink(inputPath).catch(() => {})
		if (outputPath) await fs.unlink(outputPath).catch(() => {})
	}
})

// ==========================================
// 2. XLSX → Images (Har bir sheet alohida rasm)
// ==========================================
app.post('/api/convert/xlsx-to-images', upload.single('file'), async (req, res) => {
	let inputPath = null
	let outputDir = null
	let zipPath = null

	try {
		if (!req.file) {
			return res.status(400).json({ error: 'Fayl yuklanmadi' })
		}

		inputPath = req.file.path
		const baseName = path.basename(inputPath, path.extname(inputPath))
		outputDir = path.join('/tmp/conversions', `${baseName}-images`)
		await fs.mkdir(outputDir, { recursive: true })

		// Excel o'qish
		const workbook = XLSX.readFile(inputPath)

		// Har bir sheet uchun rasm yaratish
		for (let i = 0; i < workbook.SheetNames.length; i++) {
			const sheetName = workbook.SheetNames[i]
			const sheet = workbook.Sheets[sheetName]
			const data = XLSX.utils.sheet_to_json(sheet, { header: 1 })

			// HTML table yaratish
			let html = '<table border="1" style="border-collapse: collapse;">'
			data.forEach((row) => {
				html += '<tr>'
				row.forEach((cell) => {
					html += `<td style="padding: 8px; min-width: 100px;">${cell || ''}</td>`
				})
				html += '</tr>'
			})
			html += '</table>'

			// HTML ni rasm qilish (Puppeteer orqali)
			const puppeteer = require('puppeteer')
			const browser = await puppeteer.launch({ args: ['--no-sandbox'] })
			const page = await browser.newPage()
			await page.setContent(html)
			const imagePath = path.join(outputDir, `sheet_${i + 1}_${sheetName}.png`)
			await page.screenshot({ path: imagePath, fullPage: true })
			await browser.close()
		}

		// ZIP yaratish
		zipPath = path.join('/tmp/conversions', `${baseName}-images.zip`)
		const output = require('fs').createWriteStream(zipPath)
		const archive = archiver('zip', { zlib: { level: 9 } })

		archive.pipe(output)
		archive.directory(outputDir, false)
		await archive.finalize()

		await new Promise((resolve, reject) => {
			output.on('close', resolve)
			output.on('error', reject)
		})

		const zipBuffer = await fs.readFile(zipPath)

		res.setHeader('Content-Type', 'application/zip')
		res.setHeader('Content-Disposition', `attachment; filename="${baseName}-images.zip"`)
		res.send(zipBuffer)
	} catch (error) {
		console.error('XLSX → Images xato:', error)
		res.status(500).json({ error: 'Konvertatsiya xatosi', details: error.message })
	} finally {
		if (inputPath) await fs.unlink(inputPath).catch(() => {})
		if (outputDir) await fs.rm(outputDir, { recursive: true, force: true }).catch(() => {})
		if (zipPath) await fs.unlink(zipPath).catch(() => {})
	}
})

// ==========================================
// 3. XLSX ma'lum sheetlarini PDF qilish
// ==========================================
app.post('/api/convert/xlsx-sheets-to-pdf', upload.single('file'), async (req, res) => {
	let inputPath = null
	let outputPath = null

	try {
		if (!req.file) {
			return res.status(400).json({ error: 'Fayl yuklanmadi' })
		}

		const { sheets = '0' } = req.body // "0,1,2" format
		const selectedSheets = sheets.split(',').map((s) => parseInt(s.trim()))

		inputPath = req.file.path
		outputPath = path.join('/tmp/conversions', `${Date.now()}.pdf`)

		const workbook = XLSX.readFile(inputPath)
		const doc = new PDFDocument({ margin: 30 })
		const stream = require('fs').createWriteStream(outputPath)
		doc.pipe(stream)

		// Tanlangan sheetlar
		selectedSheets.forEach((sheetIndex, idx) => {
			if (sheetIndex >= workbook.SheetNames.length) return

			const sheetName = workbook.SheetNames[sheetIndex]
			const sheet = workbook.Sheets[sheetName]
			const data = XLSX.utils.sheet_to_json(sheet, { header: 1 })

			if (idx > 0) doc.addPage()

			// Sheet nomi
			doc.fontSize(16).text(`Sheet: ${sheetName}`, { align: 'center' })
			doc.moveDown()

			// Jadval
			const cellWidth = 80
			const cellHeight = 25
			let yPos = doc.y

			data.forEach((row, rowIndex) => {
				let xPos = 30
				row.forEach((cell) => {
					doc.rect(xPos, yPos, cellWidth, cellHeight).stroke()
					doc.fontSize(10).text(String(cell || ''), xPos + 5, yPos + 8, {
						width: cellWidth - 10,
						ellipsis: true,
					})
					xPos += cellWidth
				})
				yPos += cellHeight

				if (yPos > 700 && rowIndex < data.length - 1) {
					doc.addPage()
					yPos = 50
				}
			})
		})

		doc.end()

		await new Promise((resolve, reject) => {
			stream.on('finish', resolve)
			stream.on('error', reject)
		})

		const pdfBuffer = await fs.readFile(outputPath)

		res.setHeader('Content-Type', 'application/pdf')
		res.setHeader('Content-Disposition', `attachment; filename="sheets-${Date.now()}.pdf"`)
		res.send(pdfBuffer)
	} catch (error) {
		console.error('XLSX Sheets → PDF xato:', error)
		res.status(500).json({ error: 'Konvertatsiya xatosi', details: error.message })
	} finally {
		if (inputPath) await fs.unlink(inputPath).catch(() => {})
		if (outputPath) await fs.unlink(outputPath).catch(() => {})
	}
})



console.log('✅ XLSX Converter API endpoints qo\'shildi')

// API endpoint для конвертации множественных файлов
app.post('/api/convert/batch', upload.array('files', 10), async (req, res) => {
	const { operation } = req.body
	const files = req.files
	const tempFiles = []

	try {
		if (!files || files.length === 0) {
			return res.status(400).json({ error: 'Файлы не загружены' })
		}

		// Здесь можно добавить логику для batch операций
		res.json({
			message: 'Batch конвертация завершена',
			filesProcessed: files.length,
		})
	} catch (error) {
		console.error('Ошибка batch конвертации:', error)
		res.status(500).json({ error: error.message })
	} finally {
		// Очистка
		for (const file of tempFiles) {
			try {
				await fs.unlink(file).catch(() => {})
			} catch (err) {
				console.error('Ошибка очистки:', err)
			}
		}
	}
})

// Главная страница
app.get('/', (req, res) => {
	res.sendFile(path.join(__dirname, 'public', 'index.html'))
})

app.use((req, res) => {
	res.status(404).json({ error: 'Endpoint не найден' })
})

app.use((err, req, res, next) => {
	console.error('Server error:', err)
	res.status(500).json({
		error: 'Внутренняя ошибка сервера',
		details: process.env.NODE_ENV === 'development' ? err.message : undefined,
	})
})

app.listen(PORT, () => {
	console.log(`🚀 Сервер запущен на порту ${PORT}`)
	console.log(`📄 Доступен по адресу: http://localhost:${PORT}`)
})
