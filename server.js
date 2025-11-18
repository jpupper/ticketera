const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const sharp = require('sharp');
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
app.use(express.json());

// Sistema de tracking de participantes con múltiples ticketeadoras
const TICKETEADORAS = {
  'default': {
    name: 'Ticketeadora 1',
    logo: 'logo.png',
    file: 'participants.json'
  },
  'veladero': {
    name: 'Veladero',
    logo: 'veladerologo2.png',
    file: 'participants_veladero.json'
  }
};

function getParticipantsFile(ticketeadora = 'default') {
  const config = TICKETEADORAS[ticketeadora] || TICKETEADORAS['default'];
  return path.join(__dirname, config.file);
}

function getTicketeadoraConfig(ticketeadora = 'default') {
  return TICKETEADORAS[ticketeadora] || TICKETEADORAS['default'];
}

function getParticipantsData(ticketeadora = 'default') {
  const participantsFile = getParticipantsFile(ticketeadora);
  if (!fs.existsSync(participantsFile)) {
    return { lastNumber: 0, history: [] };
  }
  try {
    const data = fs.readFileSync(participantsFile, 'utf-8');
    return JSON.parse(data);
  } catch (err) {
    console.error(`Error leyendo ${participantsFile}:`, err);
    return { lastNumber: 0, history: [] };
  }
}

function saveParticipantsData(data, ticketeadora = 'default') {
  const participantsFile = getParticipantsFile(ticketeadora);
  try {
    fs.writeFileSync(participantsFile, JSON.stringify(data, null, 2), 'utf-8');
  } catch (err) {
    console.error(`Error guardando ${participantsFile}:`, err);
  }
}

function getNextParticipantNumber(ticketeadora = 'default') {
  const data = getParticipantsData(ticketeadora);
  data.lastNumber += 1;
  const now = new Date();
  const participantInfo = {
    number: data.lastNumber,
    date: now.toLocaleDateString('es-AR'),
    time: now.toLocaleTimeString('es-AR'),
    timestamp: now.toISOString()
  };
  data.history.push(participantInfo);
  saveParticipantsData(data, ticketeadora);
  return participantInfo;
}

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

// Endpoint para obtener lista de ticketeadoras
app.options('/ticketeadoras', cors());
app.get('/ticketeadoras', async (req, res) => {
  try {
    const ticketeadoras = Object.keys(TICKETEADORAS).map(key => ({
      id: key,
      name: TICKETEADORAS[key].name,
      logo: TICKETEADORAS[key].logo
    }));
    return res.json({ ok: true, ticketeadoras });
  } catch (err) {
    console.error('Error obteniendo ticketeadoras:', err);
    return res.status(500).json({ ok: false, error: err.message });
  }
});

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
app.options('/print', cors());
app.post('/print', async (req, res) => {
  try {
    const { printer, ticketeadora = 'default' } = req.body;
    
    // Obtener configuración de la ticketeadora
    const config = getTicketeadoraConfig(ticketeadora);
    
    // Obtener siguiente número de participante
    const participant = getNextParticipantNumber(ticketeadora);

    // Preparar carpeta temporal para el PDF
    const tmpDir = path.join(__dirname, 'tmp');
    fs.mkdirSync(tmpDir, { recursive: true });
    const pdfPath = path.join(tmpDir, `ticket-${Date.now()}.pdf`);

    // Crear el PDF con ancho aproximado de rollo 80mm (226pt)
    const pageWidth = 226;
    const pageHeight = 500;

    const doc = new PDFDocument({
      size: [pageWidth, pageHeight],
      margins: { top: 12, bottom: 12, left: 12, right: 12 }
    });

    const stream = fs.createWriteStream(pdfPath);
    doc.pipe(stream);

    const contentWidth = pageWidth - doc.page.margins.left - doc.page.margins.right;

    // Fecha y hora (común para ambos diseños)
    const now = new Date();
    const fecha = now.toLocaleDateString('es-AR', { 
      day: '2-digit', 
      month: '2-digit', 
      year: 'numeric' 
    });
    const hora = now.getHours().toString().padStart(2, '0');
    const minutos = now.getMinutes().toString().padStart(2, '0');
    const segundos = now.getSeconds().toString().padStart(2, '0');

    // Diseño según ticketeadora
    if (ticketeadora === 'veladero') {
      // ========== DISEÑO VELADERO ==========
      
      // Texto "CINE INMERSIVO" arriba a la izquierda
      doc.font('Helvetica-Bold').fontSize(9).text('CINE INMERSIVO', doc.page.margins.left, doc.y, { 
        align: 'left',
        width: contentWidth
      });
      doc.moveDown(0.5);
      
      // Logo centrado
      const logoPath = path.join(__dirname, 'public', config.logo);
      if (fs.existsSync(logoPath)) {
        try {
          const imgBuffer = await transformImageForThermal(logoPath);
          if (imgBuffer) {
            const logoWidth = Math.round(contentWidth * 0.90);
            const xPosition = doc.page.margins.left + (contentWidth - logoWidth) / 2;
            doc.image(imgBuffer, xPosition, doc.y, { width: logoWidth });
            doc.moveDown(10);
          }
        } catch (imgErr) {
          console.warn('Error procesando logo:', imgErr.message);
          doc.moveDown(2);
        }
      }

      // Texto "VIVI LA EXPERIENCIA" (pequeño)
      doc.font('Helvetica-Bold').fontSize(11).text('VIVI LA EXPERIENCIA', { align: 'center' });
      doc.moveDown(1.2);

      // Imagen "un futuro más brillante"
      const sloganPath = path.join(__dirname, 'public', 'futurobrillante.png');
      if (fs.existsSync(sloganPath)) {
        try {
          const sloganBuffer = await transformImageForThermal(sloganPath);
          if (sloganBuffer) {
            const sloganWidth = Math.round(contentWidth * 0.85);
            const xPositionSlogan = doc.page.margins.left + (contentWidth - sloganWidth) / 2;
            doc.image(sloganBuffer, xPositionSlogan, doc.y, { width: sloganWidth });
            doc.moveDown(8);
          }
        } catch (imgErr) {
          console.warn('Error procesando imagen slogan:', imgErr.message);
          // Fallback a texto si falla la imagen
          doc.font('Helvetica-Bold').fontSize(24).text('un futuro', { align: 'center' });
          doc.fontSize(24).text('más', { align: 'center' });
          doc.fontSize(24).text('brillante', { align: 'center' });
          doc.moveDown(1.5);
        }
      } else {
        // Fallback a texto si no existe la imagen
        doc.font('Helvetica-Bold').fontSize(24).text('un futuro', { align: 'center' });
        doc.fontSize(24).text('más', { align: 'center' });
        doc.fontSize(24).text('brillante', { align: 'center' });
        doc.moveDown(1.5);
      }

      // Separador ondulado
      doc.font('Helvetica').fontSize(10).text('~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~', { align: 'center' });
      doc.moveDown(0.5);

      // Fecha y hora
      doc.font('Helvetica-Bold').fontSize(12).text(`Fecha: ${fecha}`, { align: 'center' });
      doc.fontSize(12).text(`Hora: ${hora}:${minutos}:${segundos}`, { align: 'center' });
      doc.moveDown(0.5);
      doc.font('Helvetica').fontSize(10).text('~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~ ~', { align: 'center' });
      doc.moveDown(1);

      // Número de participante
      doc.font('Helvetica-Bold').fontSize(13).text('PARTICIPANTE', { align: 'center' });
      doc.moveDown(0.3);
      doc.font('Helvetica-Bold').fontSize(32).text(`#${participant.number}`, { align: 'center' });
      doc.moveDown(1.2);

      // Mensaje final
      doc.font('Helvetica-Bold').fontSize(11).text('GRACIAS POR PARTICIPAR', { align: 'center' });

    } else {
      // ========== DISEÑO DEFAULT (Ticketeadora 1) ==========
      
      // Marco decorativo superior
      doc.fontSize(10).text('==============================', { align: 'center' });
      doc.fontSize(10).text('*  *  *  *  *  *  *  *  *  *', { align: 'center' });
      doc.moveDown(0.8);

      // Logo centrado
      const logoPath = path.join(__dirname, 'public', config.logo);
      if (fs.existsSync(logoPath)) {
        try {
          const imgBuffer = await transformImageForThermal(logoPath);
          if (imgBuffer) {
            const logoWidth = Math.round(contentWidth * 0.85);
            const xPosition = doc.page.margins.left + (contentWidth - logoWidth) / 2;
            doc.image(imgBuffer, xPosition, doc.y, { width: logoWidth });
            doc.moveDown(8);
          } else {
            doc.moveDown(2);
          }
        } catch (imgErr) {
          console.warn('Error procesando logo:', imgErr.message);
          doc.moveDown(2);
        }
      } else {
        doc.moveDown(2);
      }

      // Fecha y hora
      doc.font('Helvetica-Bold').fontSize(11).text('==============================', { align: 'center' });
      doc.moveDown(0.5);
      doc.font('Helvetica-Bold').fontSize(13).text(`Fecha: ${fecha}`, { align: 'center' });
      doc.fontSize(13).text(`Hora: ${hora}:${minutos}:${segundos}`, { align: 'center' });
      doc.moveDown(0.7);
      doc.font('Helvetica-Bold').fontSize(11).text('==============================', { align: 'center' });
      doc.moveDown(1);

      // Número de participante
      doc.font('Helvetica-Bold').fontSize(14).text('PARTICIPANTE', { align: 'center' });
      doc.moveDown(0.3);
      doc.font('Helvetica-Bold').fontSize(36).text(`#${participant.number}`, { align: 'center' });
      doc.moveDown(1);

      // Marco decorativo inferior
      doc.fontSize(10).text('*  *  *  *  *  *  *  *  *  *', { align: 'center' });
      doc.fontSize(10).text('==============================', { align: 'center' });
      doc.moveDown(0.8);
      doc.font('Helvetica-Bold').fontSize(11).text('GRACIAS POR PARTICIPAR', { align: 'center' });
    }

    doc.end();

    await new Promise((resolve, reject) => {
      stream.on('finish', resolve);
      stream.on('error', reject);
    });

    // Validar impresora seleccionada si se proporcionó y enviar a la impresora
    const printOptions = {};
    if (printer) {
      printOptions.printer = printer;
      if (typeof getPrinters === 'function') {
        try {
          const list = await getPrinters();
          const names = Array.isArray(list)
            ? list.map(resolvePrinterName).filter(Boolean)
            : [];
          const lower = names.map(n => n.toLowerCase());
          if (!lower.includes(printer.toLowerCase())) {
            return res.status(400).json({ ok: false, error: 'La impresora seleccionada no está disponible.' });
          }
        } catch (_) {
          // Si falla el listado, continuamos e intentamos imprimir igualmente
        }
      }
    }

    await print(pdfPath, printOptions);

    // Eliminar archivo temporal
    try { fs.unlinkSync(pdfPath); } catch {}

    return res.json({ 
      ok: true, 
      message: 'Ticket enviado a impresión.',
      participantNumber: participant.number
    });
  } catch (err) {
    console.error('Error en impresión:', err);
    return res.status(500).json({ ok: false, error: 'Falló la impresión: ' + err.message });
  }
});

// Endpoint para reiniciar el contador de participantes
app.options('/reset', cors());
app.post('/reset', async (req, res) => {
  try {
    const { ticketeadora = 'default' } = req.body;
    const data = getParticipantsData(ticketeadora);
    data.lastNumber = 0;
    saveParticipantsData(data, ticketeadora);
    return res.json({ ok: true, message: 'Contador reiniciado a 0' });
  } catch (err) {
    console.error('Error al reiniciar contador:', err);
    return res.status(500).json({ ok: false, error: 'Falló el reinicio: ' + err.message });
  }
});

// Endpoint para obtener el historial de participantes
app.options('/history', cors());
app.get('/history', async (req, res) => {
  try {
    const ticketeadora = req.query.ticketeadora || 'default';
    const data = getParticipantsData(ticketeadora);
    return res.json({ 
      ok: true, 
      lastNumber: data.lastNumber,
      history: data.history || [] 
    });
  } catch (err) {
    console.error('Error al obtener historial:', err);
    return res.status(500).json({ ok: false, error: 'Falló obtener historial: ' + err.message });
  }
});

const PORT = process.env.PORT || 5450;
app.listen(PORT, () => {
  console.log(`Servidor arrancado en http://localhost:${PORT}/`);
});