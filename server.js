const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const sharp = require('sharp');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const { print, getPrinters, getDefaultPrinter } = require('pdf-to-printer');
const path = require('path');
const fs = require('fs');

const app = express();

// Habilitar CORS para permitir peticiones desde otros orígenes (p. ej., Apache)
app.use(cors());
// Express 5 ya no admite el comodín '*' en rutas.
// Para preflight, declaramos OPTIONS sólo en endpoints específicos más abajo.

// Servir archivos estáticos (frontend)
app.use(express.static(path.join(__dirname, 'public')));

// Configurar almacenamiento de imágenes con multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = path.join(__dirname, 'uploads');
    fs.mkdirSync(dir, { recursive: true });
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || '.png';
    cb(null, `imagen-${Date.now()}${ext}`);
  }
});
const upload = multer({ storage });

// Utilidad para resolver el nombre de la impresora desde diferentes formatos
function resolvePrinterName(p) {
  if (!p) return null;
  if (typeof p === 'string') return p;
  if (typeof p !== 'object') return null;
  const candidateKeys = [
    'name','Name','printerName','PrinterName','deviceName','DeviceName','deviceId','DeviceId','DeviceID','Printer','PRINTER'
  ];
  for (const key of candidateKeys) {
    const val = p[key];
    if (typeof val === 'string' && val.trim()) return val.trim();
  }
  // Fallback: primera propiedad string no vacía
  for (const k of Object.keys(p)) {
    const v = p[k];
    if (typeof v === 'string' && v.trim()) return v.trim();
  }
  return null;
}

function execCmd(cmd) {
  return new Promise((resolve, reject) => {
    exec(cmd, { windowsHide: true }, (err, stdout, stderr) => {
      if (err) return reject(err);
      resolve({ stdout: stdout || '', stderr: stderr || '' });
    });
  });
}

async function listPrintersFallback() {
  // Intento 1: PowerShell Get-Printer (puede no estar disponible en todas las versiones)
  try {
    const { stdout } = await execCmd('powershell -NoProfile -Command "Get-Printer | Select-Object -ExpandProperty Name"');
    const names = stdout
      .split(/\r?\n/)
      .map(s => s.trim())
      .filter(Boolean);
    if (names.length) return names;
  } catch (_) {}

  // Intento 2: WMI vía PowerShell (más compatible)
  try {
    const { stdout } = await execCmd('powershell -NoProfile -Command "Get-WmiObject -Class Win32_Printer | Select-Object -ExpandProperty Name"');
    const names = stdout
      .split(/\r?\n/)
      .map(s => s.trim())
      .filter(Boolean);
    if (names.length) return names;
  } catch (_) {}

  // Intento 3: WMIC (puede estar deprecado en versiones modernas)
  try {
    const { stdout } = await execCmd('wmic printer get Name');
    const lines = stdout.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const names = lines.filter(l => l.toLowerCase() !== 'name');
    if (names.length) return names;
  } catch (_) {}

  return [];
}

async function defaultPrinterFallback() {
  // PowerShell WMI: buscar impresora por Default=true
  try {
    const { stdout } = await execCmd('powershell -NoProfile -Command "(Get-WmiObject -Class Win32_Printer | Where-Object {$_.Default -eq $true} | Select-Object -ExpandProperty Name)"');
    const name = (stdout || '').trim();
    if (name) return name;
  } catch (_) {}

  // WMIC: obtener tabla de Name,Default y seleccionar la que tenga TRUE
  try {
    const { stdout } = await execCmd('wmic printer get Name,Default');
    const lines = stdout.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    for (const line of lines) {
      const parts = line.split(/\s{2,}/); // separar por múltiples espacios
      if (parts.length >= 2) {
        const [name, def] = parts;
        if ((def || '').toLowerCase().includes('true')) {
          return (name || '').trim();
        }
      }
    }
  } catch (_) {}

  return null;
}

async function transformImageForThermal(imagePath) {
  try {
    // Convertir a escala de grises y aplicar umbral para simular salida térmica
    const buf = await sharp(imagePath)
      .grayscale()
      .threshold(180)
      .png()
      .toBuffer();
    return buf;
  } catch (e) {
    console.warn('No se pudo transformar imagen, se usará original:', e.message);
    // Si falla la transformación, devolver el archivo original como buffer
    try {
      return await sharp(imagePath).png().toBuffer();
    } catch (_) {
      return null;
    }
  }
}

// Listar impresoras disponibles y la predeterminada
// Preflight para /printers
app.options('/printers', cors());
app.get('/printers', async (req, res) => {
  try {
    let list = [];
    let defaultPrinter = null;
    if (typeof getPrinters === 'function') {
      try {
        list = await getPrinters();
      } catch (e) {
        console.warn('getPrinters falló, usando fallback:', e.message);
        list = await listPrintersFallback();
      }
    } else {
      list = await listPrintersFallback();
    }
    if (typeof getDefaultPrinter === 'function') {
      try {
        defaultPrinter = await getDefaultPrinter();
      } catch (e) {
        console.warn('getDefaultPrinter falló, usando fallback:', e.message);
        defaultPrinter = await defaultPrinterFallback();
      }
    } else {
      defaultPrinter = await defaultPrinterFallback();
    }

    const names = Array.isArray(list)
      ? list.map(resolvePrinterName).filter(Boolean)
      : [];
    const defaultName = resolvePrinterName(defaultPrinter);

    res.json({ ok: true, printers: names, defaultPrinter: defaultName });
  } catch (err) {
    // No lanzar error duro; retornar lista vacía para que el frontend pueda seguir operando
    console.error('Error inesperado listando impresoras:', err);
    res.json({ ok: true, printers: [], defaultPrinter: null });
  }
});

// Endpoint principal para imprimir
// Preflight para /print
app.options('/print', cors());
app.post('/print', upload.single('image'), async (req, res) => {
  try {
    const title = (req.body.title || '').trim();
    const description = (req.body.description || '').trim();
    const imagePath = req.file ? req.file.path : null;
    const selectedPrinter = (req.body.printer || '').trim();

    if (!title || !description) {
      return res.status(400).json({ ok: false, error: 'Título y descripción son requeridos.' });
    }

    // Preparar carpeta temporal para el PDF
    const tmpDir = path.join(__dirname, 'tmp');
    fs.mkdirSync(tmpDir, { recursive: true });
    const pdfPath = path.join(tmpDir, `ticket-${Date.now()}.pdf`);

    // Crear el PDF con ancho aproximado de rollo 80mm (226pt)
    const pageWidth = 226; // puntos
    const pageHeight = 800; // altura amplia para evitar cortes

    const doc = new PDFDocument({
      size: [pageWidth, pageHeight],
      margins: { top: 12, bottom: 12, left: 12, right: 12 }
    });

    const stream = fs.createWriteStream(pdfPath);
    doc.pipe(stream);

    const contentWidth = pageWidth - doc.page.margins.left - doc.page.margins.right;
    const imageWidth = Math.round(contentWidth * 0.35); // imagen mucho más chica

    // Título
    doc.font('Helvetica-Bold').fontSize(18).text(title, { align: 'center', width: contentWidth });
    doc.moveDown(0.5);

    // Descripción
    doc.font('Helvetica').fontSize(12).text(description, { align: 'left', width: contentWidth });
    doc.moveDown(0.5);

    // Imagen (si se subió) transformada para impresora térmica
    if (imagePath) {
      try {
        const imgBuffer = await transformImageForThermal(imagePath);
        if (imgBuffer) {
          doc.image(imgBuffer, { width: imageWidth, align: 'center' });
          doc.moveDown(0.5);
        }
      } catch (imgErr) {
        console.warn('No se pudo insertar imagen en el PDF:', imgErr.message);
      }
    }

    // Separador opcional
    doc.moveDown(0.5);
    doc.fontSize(10).text('------------------------------', { align: 'center' });

    doc.end();

    await new Promise((resolve, reject) => {
      stream.on('finish', resolve);
      stream.on('error', reject);
    });

    // Validar impresora seleccionada si se proporcionó y enviar a la impresora
    const printOptions = {};
    if (selectedPrinter) {
      printOptions.printer = selectedPrinter;
      if (typeof getPrinters === 'function') {
        try {
          const list = await getPrinters();
          const names = Array.isArray(list)
            ? list.map(resolvePrinterName).filter(Boolean)
            : [];
          const lower = names.map(n => n.toLowerCase());
          if (!lower.includes(selectedPrinter.toLowerCase())) {
            return res.status(400).json({ ok: false, error: 'La impresora seleccionada no está disponible.' });
          }
        } catch (_) {
          // Si falla el listado, continuamos e intentamos imprimir igualmente
        }
      }
    }

    await print(pdfPath, printOptions);

    // Eliminar archivos temporales
    try { fs.unlinkSync(pdfPath); } catch {}
    if (imagePath) { try { fs.unlinkSync(imagePath); } catch {} }

    return res.json({ ok: true, message: 'Ticket enviado a impresión.' });
  } catch (err) {
    console.error('Error en impresión:', err);
    return res.status(500).json({ ok: false, error: 'Falló la impresión: ' + err.message });
  }
});

// Endpoint de previsualización: genera el PDF y lo devuelve en la respuesta
// Preflight para /preview
app.options('/preview', cors());
app.post('/preview', upload.single('image'), async (req, res) => {
  try {
    const title = (req.body.title || '').trim();
    const description = (req.body.description || '').trim();
    const imagePath = req.file ? req.file.path : null;

    if (!title || !description) {
      return res.status(400).json({ ok: false, error: 'Título y descripción son requeridos.' });
    }

    const pageWidth = 226;
    const pageHeight = 800;
    const doc = new PDFDocument({ size: [pageWidth, pageHeight], margins: { top: 12, bottom: 12, left: 12, right: 12 } });
    const contentWidth = pageWidth - doc.page.margins.left - doc.page.margins.right;
    const imageWidth = Math.round(contentWidth * 0.35);

    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => {
      const pdfBuffer = Buffer.concat(chunks);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'inline; filename="preview.pdf"');
      res.send(pdfBuffer);
    });

    // Contenido
    doc.font('Helvetica-Bold').fontSize(18).text(title, { align: 'center', width: contentWidth });
    doc.moveDown(0.5);
    doc.font('Helvetica').fontSize(12).text(description, { align: 'left', width: contentWidth });
    doc.moveDown(0.5);

    if (imagePath) {
      try {
        const imgBuffer = await transformImageForThermal(imagePath);
        if (imgBuffer) {
          doc.image(imgBuffer, { width: imageWidth, align: 'center' });
          doc.moveDown(0.5);
        }
      } catch (imgErr) {
        console.warn('No se pudo insertar imagen en el PDF de preview:', imgErr.message);
      }
    }

    doc.moveDown(0.5);
    doc.fontSize(10).text('------------------------------', { align: 'center' });

    doc.end();

    // Limpiar imagen temporal
    if (imagePath) { try { fs.unlinkSync(imagePath); } catch (_) {} }
  } catch (err) {
    console.error('Error en preview:', err);
    return res.status(500).json({ ok: false, error: 'Falló la previsualización: ' + err.message });
  }
});

const PORT = process.env.PORT || 5450;
app.listen(PORT, () => {
  console.log(`Servidor arrancado en http://localhost:${PORT}/`);
});