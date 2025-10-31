const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const { print, getPrinters, getDefaultPrinter } = require('pdf-to-printer');
const path = require('path');
const fs = require('fs');

const app = express();

// Habilitar CORS para permitir peticiones desde otros orígenes (p. ej., Apache)
app.use(cors());

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

// Listar impresoras disponibles y la predeterminada
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

    // Título
    doc.font('Helvetica-Bold').fontSize(18).text(title, { align: 'center', width: contentWidth });
    doc.moveDown(0.5);

    // Descripción
    doc.font('Helvetica').fontSize(12).text(description, { align: 'left', width: contentWidth });
    doc.moveDown(0.5);

    // Imagen (si se subió)
    if (imagePath) {
      try {
        doc.image(imagePath, { width: contentWidth, align: 'center' });
        doc.moveDown(0.5);
      } catch (imgErr) {
        // Si la imagen no se puede leer, continuar sin ella
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

const PORT = process.env.PORT || 5450;
app.listen(PORT, () => {
  console.log(`Servidor arrancado en http://localhost:${PORT}/`);
});