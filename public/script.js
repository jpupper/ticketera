   const statusEl = document.getElementById('status');
const printerSelect = document.getElementById('printer');
const printBtn = document.getElementById('printBtn');
const previewDate = document.getElementById('previewDate');
const previewTime = document.getElementById('previewTime');
const previewNumber = document.getElementById('previewNumber');
let apiBase = window.location.origin;

function updatePreview() {
  const now = new Date();
  const fecha = now.toLocaleDateString('es-AR', { 
    day: '2-digit', 
    month: '2-digit', 
    year: 'numeric' 
  });
  const hora = now.getHours().toString().padStart(2, '0');
  const minutos = now.getMinutes().toString().padStart(2, '0');
  const segundos = now.getSeconds().toString().padStart(2, '0');
  
  previewDate.textContent = fecha;
  previewTime.textContent = `${hora}:${minutos}:${segundos}`;
}

// Actualizar preview cada segundo
setInterval(updatePreview, 1000);
updatePreview();

async function generateTicket() {
  statusEl.textContent = '';
  printBtn.disabled = true;
  try {
    const printer = printerSelect.value;
    const resp = await fetch(apiBase + '/print', { 
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ printer })
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      throw new Error(data.error || 'Error al imprimir');
    }
    statusEl.textContent = `¡Ticket #${data.participantNumber} generado!`;
    statusEl.style.color = '#198754';
    previewNumber.textContent = data.participantNumber + 1;
  } catch (err) {
    statusEl.textContent = 'Falló la impresión: ' + err.message;
    statusEl.style.color = '#dc3545';
  } finally {
    printBtn.disabled = false;
  }
}

printBtn.addEventListener('click', generateTicket);

// Listener para tecla U
document.addEventListener('keydown', (e) => {
  if (e.key.toLowerCase() === 'u' && !printBtn.disabled) {
    generateTicket();
  }
});

async function loadPrinters() {
  try {
    let resp = await fetch(apiBase + '/printers');
    let ct = resp.headers.get('content-type') || '';
    if (!ct.includes('application/json')) {
      resp = await fetch('http://localhost:5450/printers');
      apiBase = 'http://localhost:5450';
      ct = resp.headers.get('content-type') || '';
    }
    const data = await resp.json();
    if (!resp.ok || !data.ok) throw new Error(data.error || 'No se pudieron listar impresoras');
    let { printers, defaultPrinter } = data;
    if (!Array.isArray(printers)) printers = [];
    printers.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      printerSelect.appendChild(opt);
    });
    const defName = typeof defaultPrinter === 'string' ? defaultPrinter : (defaultPrinter && (defaultPrinter.name || defaultPrinter.PrinterName || defaultPrinter.deviceId || defaultPrinter.DeviceID));
    if (defName) {
      const found = Array.from(printerSelect.options).find(o => o.value === defName);
      if (found) printerSelect.value = defName;
    }
  } catch (err) {
    statusEl.textContent = 'No se pudo cargar la lista de impresoras: ' + err.message;
    statusEl.style.color = '#dc3545';
  }
}

window.addEventListener('DOMContentLoaded', loadPrinters);