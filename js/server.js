const express = require('express')
const multer = require('multer')
const { execFile } = require('child_process')
const { promisify } = require('util')
const fs = require('fs').promises
const path = require('path')
const { v4: uuidv4 } = require('uuid')

const execFileAsync = promisify(execFile)

const app = express()
const PORT = process.env.PORT || 3000
const TEMP_DIR = '/tmp/conversions'
const TIMEOUT_MS = 120000 // 2 minutes

const SUPPORTED_INPUT_FORMATS = [
	'doc',
	'docx',
	'odt',
	'rtf',
	'txt',
	'ppt',
	'pptx',
	'odp',
	'xls',
	'xlsx',
	'ods',
	'csv',
	'html',
	'htm',
	'xml',
	'wpd',
	'wps', // условная поддержка
]

const SUPPORTED_OUTPUT_FORMATS = {
	pdf: 'pdf',
	docx: 'docx',
	odt: 'odt',
	txt: 'txt:Text',
	html: 'html',
	rtf: 'rtf',
	xlsx: 'xlsx',
	ods: 'ods',
	pptx: 'pptx',
	odp: 'odp',
}

// Configure multer for file uploads
const storage = multer.diskStorage({
	destination: async (req, file, cb) => {
		const uploadDir = path.join(TEMP_DIR, uuidv4())
		await fs.mkdir(uploadDir, { recursive: true })
		req.uploadDir = uploadDir
		cb(null, uploadDir)
	},
	filename: (req, file, cb) => {
		cb(null, file.originalname)
	},
})

const upload = multer({
	storage,
	limits: {
		fileSize: 50 * 1024 * 1024, // 50MB
	},
	fileFilter: (req, file, cb) => {
		const ext = path.extname(file.originalname).toLowerCase().slice(1)
		if (SUPPORTED_INPUT_FORMATS.includes(ext)) {
			cb(null, true)
		} else {
			cb(new Error(`Unsupported file format: ${ext}`))
		}
	},
})

// Middleware
app.use(express.json())
app.use(express.static('public'))

// Health check endpoint
app.get('/health', (req, res) => {
	res.json({ status: 'ok', timestamp: new Date().toISOString() })
})

// Get supported formats
app.get('/api/formats', (req, res) => {
	res.json({
		input: SUPPORTED_INPUT_FORMATS,
		output: Object.keys(SUPPORTED_OUTPUT_FORMATS),
	})
})

// Convert file endpoint
app.post('/api/convert', upload.single('file'), async (req, res) => {
	let uploadDir = req.uploadDir

	try {
		if (!req.file) {
			return res.status(400).json({ error: 'No file uploaded' })
		}

		const outputFormat = req.body.format || 'pdf'

		if (!SUPPORTED_OUTPUT_FORMATS[outputFormat]) {
			return res.status(400).json({
				error: `Unsupported output format: ${outputFormat}`,
			})
		}

		const inputPath = req.file.path
		const inputExt = path.extname(req.file.originalname).toLowerCase().slice(1)
		const baseName = path.basename(
			req.file.originalname,
			path.extname(req.file.originalname)
		)

		console.log(
			`Converting ${req.file.originalname} (${inputExt}) to ${outputFormat}`
		)

		// Run LibreOffice conversion
		const convertFormat = SUPPORTED_OUTPUT_FORMATS[outputFormat]

		try {
			await execFileAsync(
				'libreoffice',
				[
					'--headless',
					'--convert-to',
					convertFormat,
					'--outdir',
					uploadDir,
					inputPath,
				],
				{
					timeout: TIMEOUT_MS,
					maxBuffer: 10 * 1024 * 1024, // 10MB buffer
				}
			)
		} catch (execError) {
			console.error('LibreOffice conversion error:', execError)
			throw new Error(`Conversion failed: ${execError.message}`)
		}

		// Find the converted file
		const files = await fs.readdir(uploadDir)
		const outputFile = files.find(
			f =>
				f !== req.file.originalname &&
				f.startsWith(baseName) &&
				f.endsWith(`.${outputFormat}`)
		)

		if (!outputFile) {
			throw new Error('Conversion completed but output file not found')
		}

		const outputPath = path.join(uploadDir, outputFile)

		// Check if file exists and has content
		const stats = await fs.stat(outputPath)
		if (stats.size === 0) {
			throw new Error('Conversion produced empty file')
		}

		console.log(`Conversion successful: ${outputFile} (${stats.size} bytes)`)

		// Send the converted file
		res.download(outputPath, `${baseName}.${outputFormat}`, async err => {
			// Cleanup after download completes (or fails)
			try {
				await fs.rm(uploadDir, { recursive: true, force: true })
			} catch (cleanupError) {
				console.error('Cleanup error:', cleanupError)
			}

			if (err) {
				console.error('Download error:', err)
			}
		})
	} catch (error) {
		console.error('Conversion error:', error)

		// Cleanup on error
		if (uploadDir) {
			try {
				await fs.rm(uploadDir, { recursive: true, force: true })
			} catch (cleanupError) {
				console.error('Cleanup error:', cleanupError)
			}
		}

		res.status(500).json({
			error: error.message || 'Conversion failed',
		})
	}
})

// Error handling middleware
app.use((error, req, res, next) => {
	if (error instanceof multer.MulterError) {
		if (error.code === 'LIMIT_FILE_SIZE') {
			return res.status(400).json({ error: 'File too large (max 50MB)' })
		}
		return res.status(400).json({ error: error.message })
	}

	if (error) {
		return res.status(400).json({ error: error.message })
	}

	next()
})

// Periodic cleanup of old temp files (safety net)
setInterval(async () => {
	try {
		const dirs = await fs.readdir(TEMP_DIR)
		const now = Date.now()

		for (const dir of dirs) {
			const dirPath = path.join(TEMP_DIR, dir)
			try {
				const stats = await fs.stat(dirPath)
				// Remove directories older than 1 hour
				if (now - stats.mtimeMs > 3600000) {
					await fs.rm(dirPath, { recursive: true, force: true })
					console.log(`Cleaned up old directory: ${dir}`)
				}
			} catch (err) {
				// Ignore errors for individual directories
			}
		}
	} catch (err) {
		console.error('Cleanup task error:', err)
	}
}, 300000) // Run every 5 minutes

// Start server
app.listen(PORT, () => {
	console.log(`LibreOffice Converter running on port ${PORT}`)
	console.log(`Supported input formats: ${SUPPORTED_INPUT_FORMATS.join(', ')}`)
	console.log(
		`Supported output formats: ${Object.keys(SUPPORTED_OUTPUT_FORMATS).join(
			', '
		)}`
	)
})
